VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCadAuxGovernancas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Governanças"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10620
   Icon            =   "frmCadAuxGovernancas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.Toolbar tlbCadastroAuxGovernancas 
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
   Begin VB.Frame fraListaCadAuxGovernancas 
      Caption         =   "Lista de governanças cadastradas"
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
               Picture         =   "frmCadAuxGovernancas.frx":0442
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAuxGovernancas.frx":0554
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAuxGovernancas.frx":0666
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAuxGovernancas.frx":07D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAuxGovernancas.frx":08E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAuxGovernancas.frx":0D34
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvwCadastroAuxGovernancas 
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
   Begin VB.Frame fraDadosCadAuxGovernancas 
      Caption         =   "Dados da governança"
      Height          =   2835
      Left            =   15
      TabIndex        =   4
      Top             =   540
      Width           =   10575
      Begin VB.Frame Frame1 
         Height          =   870
         Left            =   120
         TabIndex        =   6
         Top             =   1830
         Width           =   10335
         Begin VB.CommandButton cmdGravar 
            Caption         =   "Gravar"
            Height          =   375
            Left            =   7875
            TabIndex        =   8
            Top             =   300
            Width           =   1020
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   9105
            TabIndex        =   7
            Top             =   300
            Width           =   1020
         End
      End
      Begin VB.TextBox txtDescricaoGovernanca 
         Height          =   345
         Left            =   2925
         MaxLength       =   50
         TabIndex        =   1
         Top             =   495
         Width           =   6675
      End
      Begin VB.Label lblDescricaoGovernanca 
         AutoSize        =   -1  'True
         Caption         =   "Descrição da Governança:"
         Height          =   195
         Left            =   945
         TabIndex        =   5
         Top             =   570
         Width           =   1920
      End
   End
End
Attribute VB_Name = "frmCadAuxGovernancas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub mDesabilitarBotoesToolBarGovernancas()

    With tlbCadastroAuxGovernancas
    
        .Buttons(1).Enabled = False
        .Buttons(2).Enabled = False
        .Buttons(3).Enabled = False
        
    End With
    
End Sub

Private Sub mHabilitarBotoesToolBarGovernancas()

    With tlbCadastroAuxGovernancas
    
        .Buttons(1).Enabled = True
        .Buttons(2).Enabled = True
        .Buttons(3).Enabled = True
        
    End With
    
End Sub

Private Sub mAlterarGovernanca()

    Dim strSql            As String
    Dim lngLinhasAfetadas As Long
    
    On Error GoTo ErrmAlterarGovernanca
    
    Screen.MousePointer = vbHourglass
    
    strSql = "SP_ALTERAR_GOVERNANCA " & CInt(lvwCadastroAuxGovernancas.SelectedItem.Text) & ", " & _
             "'" & txtDescricaoGovernanca.Text & "'"
             
    dbConexaoSMR.Execute strSql, lngLinhasAfetadas
    
    If lngLinhasAfetadas > 0 Then
        MsgBox "Governança alterada com sucesso ! ", vbInformation + vbOKOnly, App.Title
    End If
    
    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadAuxGovernancas.Visible = False
    fraListaCadAuxGovernancas.Visible = True
    
    'Alterando os dados na lista
    With lvwCadastroAuxGovernancas.SelectedItem
        .SubItems(1) = txtDescricaoGovernanca.Text
        .EnsureVisible
    End With
        
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrmAlterarGovernanca:

    Screen.MousePointer = vbNormal
    
    MsgBox "Ocorreu o seguinte erro na rotina mAlterarGovernanca: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub

Private Sub mDadosGovernanca()

    On Error GoTo ErrmDadosGovernanca
    
    With lvwCadastroAuxGovernancas.SelectedItem
        txtDescricaoGovernanca.Text = .SubItems(1)
    End With
    
    Exit Sub
    
ErrmDadosGovernanca:

    MsgBox "Ocorreu o seguinte erro na rotina mDadosGovernanca: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub



Private Sub mExcluirGovernanca()

    Dim strSql            As String
    Dim lngLinhasAfetadas As Long
    
    On Error GoTo ErrmExcluirGovernanca
    
    Screen.MousePointer = vbHourglass
    
    strSql = "SP_EXCLUIR_GOVERNANCA " & CInt(lvwCadastroAuxGovernancas.SelectedItem.Text)
             
    dbConexaoSMR.Execute strSql, lngLinhasAfetadas
    
    If lngLinhasAfetadas > 0 Then
        MsgBox "Governança excluída com sucesso ! ", vbInformation + vbOKOnly, App.Title
    End If
    
    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadAuxGovernancas.Visible = False
    fraListaCadAuxGovernancas.Visible = True
    
    'Preencher lista
    Call mPreencheListaGovernancas
    
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrmExcluirGovernanca:

    Screen.MousePointer = vbNormal
    
    MsgBox "Ocorreu o seguinte erro na rotina mExcluirGovernanca: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub

Private Sub mHabilitarDesabilitarCamposCadastroGovernanca(ByVal pAcao As String)

    fraDadosCadAuxGovernancas.Tag = pAcao

    If UCase(Trim(pAcao)) = "I" Then
        
        txtDescricaoGovernanca.Enabled = True
        
        txtDescricaoGovernanca.SetFocus
    
    ElseIf UCase(Trim(pAcao)) = "A" Then
        
        txtDescricaoGovernanca.Enabled = True
        
    End If
    
End Sub

Private Sub mIncluirGovernanca()

    Dim strSql                  As String
    Dim lngCodigoGovernanca     As Long
    Dim lngLinhasAfetadas       As Long
    Dim itmGovernancaInserida   As ListItem
    Dim cmdGovernanca           As ADODB.Command
    
    On Error GoTo ErrmIncluirGovernanca
    
    Screen.MousePointer = vbHourglass
    
    Set cmdGovernanca = New ADODB.Command
    
    With cmdGovernanca
    
        .ActiveConnection = dbConexaoSMR
        .CommandType = adCmdStoredProc
        .CommandText = "SP_INCLUIR_GOVERNANCA"
        
        .Parameters.Refresh
        .Parameters(1).Value = Trim(txtDescricaoGovernanca.Text)
        
        .Execute
        
        lngCodigoGovernanca = .Parameters(2).Value
        
    End With
        
    If lngCodigoGovernanca > 0 Then
        MsgBox "Governança incluída com sucesso ! ", vbInformation + vbOKOnly, App.Title
    End If
    
    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadAuxGovernancas.Visible = False
    fraListaCadAuxGovernancas.Visible = True
    
    'Adicionando na lista
    Set itmGovernancaInserida = lvwCadastroAuxGovernancas.ListItems.Add(, , Format(lngCodigoGovernanca, "000"))
    itmGovernancaInserida.SubItems(1) = txtDescricaoGovernanca.Text
    
    itmGovernancaInserida.Selected = True
    itmGovernancaInserida.EnsureVisible
        
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrmIncluirGovernanca:

    Screen.MousePointer = vbNormal
    
    MsgBox "Ocorreu o seguinte erro na rotina mIncluirGovernanca: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub

Private Sub mLimparDadosGovernanca()

    txtDescricaoGovernanca.Text = ""
    
End Sub


Private Sub mMontaCabecalhoListaGovernanca()

    'Criando os cabeçalhos das colunas da lista
    With lvwCadastroAuxGovernancas
        .ColumnHeaders.Add , , "Código", 1000
        .ColumnHeaders.Add , , "Descrição", 9370
    End With
    
End Sub

Private Sub mPreencheListaGovernancas()

    Dim rsGovernancas   As ADODB.Recordset
    Dim itmGovernanca   As ListItem
    Dim strSql          As String
    
    On Error GoTo ErrmPreencheListaGovernancas
    
    Screen.MousePointer = vbHourglass
    
    With lvwCadastroAuxGovernancas
    
        .ListItems.Clear
        .ColumnHeaders.Clear
        
        'Monta o cabeçalho da lista
        Call mMontaCabecalhoListaGovernanca
        
        strSql = "EXECUTE SP_LISTAR_GOVERNANCA"
        
        Set rsGovernancas = New ADODB.Recordset
        rsGovernancas.Open strSql, dbConexaoSMR, adOpenStatic, adLockReadOnly, adCmdText
        
        Do While Not rsGovernancas.EOF
        
            Set itmGovernanca = .ListItems.Add(, , Format(rsGovernancas.Fields("Cod_Governanca").Value, "000"))
            itmGovernanca.SubItems(1) = rsGovernancas.Fields("Desc_Governanca").Value
              
            rsGovernancas.MoveNext
            
        Loop
        
    End With
    
    rsGovernancas.Close
    Set rsGovernancas = Nothing
    
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrmPreencheListaGovernancas:

    Screen.MousePointer = vbNormal
    
    MsgBox "Ocorreu o seguinte erro na rotina mPreencheListaGovernancas: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title
           
    If rsGovernancas.State = adStateOpen Then
        rsGovernancas.Close
        Set rsGovernancas = Nothing
    End If

End Sub

Private Sub mTratarTelaAlterarGovernanca()

    On Error GoTo ErrmTratarTelaAlterarGovernanca
    
    'Desabilita os botões da ToolBar
    Call mDesabilitarBotoesToolBarGovernancas
    
    'Coloca o frame que contém os dados sobreposto ao frame de lista
    fraDadosCadAuxGovernancas.Visible = True
    fraListaCadAuxGovernancas.Visible = False
        
    Exit Sub
    
ErrmTratarTelaAlterarGovernanca:

    MsgBox "Ocorreu o seguinte erro na rotina mTratarTelaAlterarGovernanca: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub

Private Sub mTratarTelaExcluirGovernanca()

    On Error GoTo ErrmTratarTelaExcluirGovernanca
    
    'Coloca o frame que contém os dados sobreposto ao frame de lista
    fraDadosCadAuxGovernancas.Visible = True
    fraListaCadAuxGovernancas.Visible = False
    
    Exit Sub
    
ErrmTratarTelaExcluirGovernanca:

    MsgBox "Ocorreu o seguinte erro na rotina mTratarTelaExcluirGovernanca: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub


Private Sub mTratarTelaIncluirGovernanca()

    On Error GoTo ErrmTratarTelaIncluirGovernanca
    
    'Desabilita os botões da ToolBar
    Call mDesabilitarBotoesToolBarGovernancas
    
    'Limpa dados
    Call mLimparDadosGovernanca
    
    'Coloca o frame que contém os dados sobreposto ao frame de lista
    fraDadosCadAuxGovernancas.Visible = True
    fraListaCadAuxGovernancas.Visible = False
    
    Exit Sub
    
ErrmTratarTelaIncluirGovernanca:

    MsgBox "Ocorreu o seguinte erro na rotina mTratarTelaIncluirGovernanca: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title
    
End Sub

Private Function mValidarCamposGovernanca() As Boolean

    On Error GoTo ErrmValidarCamposGovernanca
    
    If Trim(txtDescricaoGovernanca.Text) = "" Then
        MsgBox "Descrição da governança não foi preenchida !", vbInformation + vbOKOnly, App.Title
        txtDescricaoGovernanca.SetFocus
        mValidarCamposGovernanca = False
        Exit Function
    End If
        
    mValidarCamposGovernanca = True
        
    Exit Function
    
ErrmValidarCamposGovernanca:

    MsgBox "Ocorreu o seguinte erro na rotina mValidarCamposGovernanca: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title
    
End Function

Private Sub cmdCancelar_Click()

    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadAuxGovernancas.Visible = False
    fraListaCadAuxGovernancas.Visible = True

    'Habilita os botões da ToolBar
    Call mHabilitarBotoesToolBarGovernancas

End Sub

Private Sub cmdGravar_Click()
    
    On Error GoTo ErrcmdGravar_Click
    
    If mValidarCamposGovernanca() Then
    
        If fraDadosCadAuxGovernancas.Tag = "I" Then
        
            'Inclusao
            Call mIncluirGovernanca
        
        ElseIf fraDadosCadAuxGovernancas.Tag = "A" Then
            
            'Alteração
            Call mAlterarGovernanca
        
        End If
        
        'Habilita os botões da ToolBar
        Call mHabilitarBotoesToolBarGovernancas
        
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
    Call mHabilitarBotoesToolBarGovernancas
        
    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadAuxGovernancas.Visible = False
    fraListaCadAuxGovernancas.Visible = True

    'Preenche a lista
    Call mPreencheListaGovernancas
    
End Sub









Private Sub Form_Unload(Cancel As Integer)

    frmPrincipal.StatusBar.Panels(1).Text = ""
    
End Sub


Private Sub lvwCadastroAuxGovernancas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    With lvwCadastroAuxGovernancas
    
        .SortKey = ColumnHeader.Index - 1
        
        If .SortOrder = lvwAscending Then
           .SortOrder = lvwDescending
        Else
           .SortOrder = lvwAscending
        End If
        
        .Sorted = True
        
    End With

End Sub


Private Sub tlbCadastroAuxGovernancas_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
    
        Case 1 'Incluir
        
            Call mTratarTelaIncluirGovernanca
            Call mHabilitarDesabilitarCamposCadastroGovernanca("I")
                    
        Case 2 'Alterar
        
            Call mTratarTelaAlterarGovernanca
            Call mDadosGovernanca
            Call mHabilitarDesabilitarCamposCadastroGovernanca("A")
                        
        Case 3 'Excluir
        
            If MsgBox("Deseja realmente excluir este registro ?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                Call mExcluirGovernanca
            End If
            
        Case 5 'Sair
            
            Unload Me
            
    End Select

End Sub


