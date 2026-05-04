VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCadAuxUsuarios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Usuários"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8505
   Icon            =   "frmCadAuxUsuarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.Toolbar tlbCadastroAuxUsuarios 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8505
      _ExtentX        =   15002
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
   Begin VB.Frame fraListaCadAuxUsuarios 
      Caption         =   "Lista de usuários cadastrados"
      Height          =   4080
      Left            =   15
      TabIndex        =   5
      Top             =   510
      Width           =   8400
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   6555
         Top             =   2010
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
               Picture         =   "frmCadAuxUsuarios.frx":0442
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAuxUsuarios.frx":0554
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAuxUsuarios.frx":0666
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAuxUsuarios.frx":07D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAuxUsuarios.frx":08E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAuxUsuarios.frx":0D34
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvwCadastroAuxUsuarios 
         Height          =   3735
         Left            =   45
         TabIndex        =   6
         Top             =   255
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   6588
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
   Begin VB.Frame fraDadosCadAuxUsuarios 
      Caption         =   "Dados do usuário"
      Height          =   4065
      Left            =   15
      TabIndex        =   7
      Top             =   525
      Width           =   8400
      Begin VB.Frame Frame1 
         Height          =   825
         Left            =   135
         TabIndex        =   17
         Top             =   3060
         Width           =   8115
         Begin VB.CommandButton cmdGravar 
            Caption         =   "Gravar"
            Height          =   375
            Left            =   5685
            TabIndex        =   19
            Top             =   285
            Width           =   1020
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   6915
            TabIndex        =   18
            Top             =   285
            Width           =   1020
         End
      End
      Begin VB.CheckBox chkCLI 
         Caption         =   "Usuário - CLI"
         Height          =   330
         Left            =   4725
         TabIndex        =   16
         Top             =   2175
         Width           =   1395
      End
      Begin VB.TextBox txtIntegrador 
         Height          =   345
         Left            =   1575
         MaxLength       =   10
         TabIndex        =   14
         Top             =   2160
         Width           =   2610
      End
      Begin VB.ComboBox cmbUnidade 
         Height          =   315
         Left            =   1560
         TabIndex        =   13
         Top             =   2610
         Width           =   3915
      End
      Begin VB.TextBox txtRedigiteSenha 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1575
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1725
         Width           =   2610
      End
      Begin VB.TextBox txtNomeUsuario 
         Height          =   345
         Left            =   1575
         MaxLength       =   50
         TabIndex        =   2
         Top             =   900
         Width           =   6675
      End
      Begin VB.TextBox txtSenha 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1575
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1320
         Width           =   2610
      End
      Begin VB.TextBox txtLogin 
         Height          =   345
         Left            =   1575
         MaxLength       =   10
         TabIndex        =   1
         Top             =   480
         Width           =   2610
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Integrador:"
         Height          =   195
         Left            =   750
         TabIndex        =   15
         Top             =   2220
         Width           =   765
      End
      Begin VB.Label lblUnidade 
         AutoSize        =   -1  'True
         Caption         =   "Unidade:"
         Height          =   195
         Index           =   1
         Left            =   855
         TabIndex        =   12
         Top             =   2655
         Width           =   630
      End
      Begin VB.Label lblRedigiteSenha 
         AutoSize        =   -1  'True
         Caption         =   "Redigite a senha:"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   11
         Top             =   1785
         Width           =   1245
      End
      Begin VB.Label lblSenha 
         AutoSize        =   -1  'True
         Caption         =   "Senha:"
         Height          =   195
         Left            =   990
         TabIndex        =   10
         Top             =   1395
         Width           =   510
      End
      Begin VB.Label lblLoginUsuario 
         AutoSize        =   -1  'True
         Caption         =   "Login:"
         Height          =   195
         Left            =   1080
         TabIndex        =   9
         Top             =   540
         Width           =   435
      End
      Begin VB.Label lblNomeUsuario 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         Height          =   195
         Left            =   1035
         TabIndex        =   8
         Top             =   975
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmCadAuxUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mAlterarUsuario()

    Dim strSql            As String
    Dim lngLinhasAfetadas As Long
    
    On Error GoTo ErrmAlterarUsuario
    
    Screen.MousePointer = vbHourglass
    
    If mValidarSenha Then
    
        If chkCLI.Value = 1 Then
            strSql = "SP_ALTERAR_USUARIO '" & txtLogin.Text & "', " & _
                     "'" & txtNomeUsuario.Text & "', " & _
                     "'" & txtSenha.Text & "', " & _
                     "'" & txtIntegrador.Text & "', " & _
                     Str(chkCLI.Value) & ", " & _
                     Trim(Str(cmbUnidade.ItemData(cmbUnidade.ListIndex))) & " "
        Else
            strSql = "SP_ALTERAR_USUARIO '" & txtLogin.Text & "', " & _
                     "'" & txtNomeUsuario.Text & "', " & _
                     "'" & txtSenha.Text & "', " & _
                     "'" & txtIntegrador.Text & "', " & _
                     Str(chkCLI.Value) & ", 0"
        End If
                 
        dbConexaoSMR.Execute strSql, lngLinhasAfetadas
        
        If lngLinhasAfetadas > 0 Then
            MsgBox "Usuário alterado com sucesso ! ", vbInformation + vbOKOnly, App.Title
        End If
        
        'Coloca o frame que contém a lista sobreposto ao frame de dados
        fraDadosCadAuxUsuarios.Visible = False
        fraListaCadAuxUsuarios.Visible = True
        
        'Alterando os dados do usuário na lista
        With lvwCadastroAuxUsuarios.SelectedItem
            .Text = txtLogin.Text
            .SubItems(1) = txtNomeUsuario.Text
            .SubItems(2) = txtSenha.Text
            .SubItems(3) = txtIntegrador.Text
            
            If chkCLI.Value = 1 Then
                .SubItems(4) = True
                .SubItems(5) = cmbUnidade.ItemData(cmbUnidade.ListIndex)
            
            Else
                .SubItems(4) = False
                .SubItems(5) = 0
            
            End If
            
            
            .EnsureVisible
        End With
        
    Else
    
        MsgBox "Senha não confere ! ", vbInformation + vbOKOnly, App.Title
        txtSenha.SetFocus
    
    End If
    
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrmAlterarUsuario:

    Screen.MousePointer = vbNormal
    
    MsgBox "Ocorreu o seguinte erro na rotina mAlterarUsuario: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub

Private Sub mDadosUsuario()
Dim I As Integer

    On Error GoTo ErrmDadosUsuario
    
    With lvwCadastroAuxUsuarios.SelectedItem
        txtLogin.Text = .Text
        txtNomeUsuario.Text = .SubItems(1)
        txtSenha.Text = .SubItems(2)
        txtIntegrador.Text = .SubItems(3)
        
        If .SubItems(4) = True Then
            chkCLI.Value = 1
            cmbUnidade.Enabled = True
            
            For I = 0 To cmbUnidade.ListCount
                If cmbUnidade.ItemData(I) = .SubItems(5) Then
                    cmbUnidade.ListIndex = I
                    Exit For
                End If
            Next
        
        Else
            chkCLI.Value = 0
            cmbUnidade.ListIndex = -1
            cmbUnidade.Enabled = False
            
        End If
        
    End With
    
    Exit Sub
    
ErrmDadosUsuario:
    
    MsgBox "Ocorreu o seguinte erro na rotina mDadosUsuario: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub



Private Sub mExcluirUsuario()

    Dim strSql            As String
    Dim lngLinhasAfetadas As Long
    
    On Error GoTo ErrmExcluirUsuario
    
    Screen.MousePointer = vbHourglass
    
    strSql = "SP_EXCLUIR_USUARIO '" & lvwCadastroAuxUsuarios.SelectedItem.Text & "'"
             
    dbConexaoSMR.Execute strSql, lngLinhasAfetadas
    
    If lngLinhasAfetadas > 0 Then
        MsgBox "Usuário excluído com sucesso ! ", vbInformation + vbOKOnly, App.Title
    End If
    
    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadAuxUsuarios.Visible = False
    fraListaCadAuxUsuarios.Visible = True
    
    'Preencher lista de usuários
    Call mPreencheListaUsuarios
    
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrmExcluirUsuario:

    Screen.MousePointer = vbNormal
    
    MsgBox "Ocorreu o seguinte erro na rotina mExcluirUsuario: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub

Private Sub mDesabilitarBotoesToolBarUsuarios()

    With tlbCadastroAuxUsuarios
    
        .Buttons(1).Enabled = False
        .Buttons(2).Enabled = False
        .Buttons(3).Enabled = False
        
    End With
    
End Sub

Private Sub mHabilitarBotoesToolBarUsuarios()

    With tlbCadastroAuxUsuarios
    
        .Buttons(1).Enabled = True
        .Buttons(2).Enabled = True
        .Buttons(3).Enabled = True
        
    End With
    
End Sub


Private Sub mHabilitarDesabilitarCamposCadastroUsuario(ByVal pAcao As String)

    fraDadosCadAuxUsuarios.Tag = pAcao

    If UCase(Trim(pAcao)) = "I" Then
        
        txtLogin.Enabled = True
        txtNomeUsuario.Enabled = True
        txtSenha.Enabled = True
        'lblRedigiteSenha.Visible = True
        txtRedigiteSenha.Visible = True
        
        txtLogin.SetFocus
    
    ElseIf UCase(Trim(pAcao)) = "A" Then
        
        txtLogin.Enabled = False
        txtNomeUsuario.Enabled = True
        txtSenha.Enabled = True
        'lblRedigiteSenha.Visible = True
        txtRedigiteSenha.Visible = True
        txtRedigiteSenha.Text = txtSenha.Text
        
    End If
    
End Sub

Private Sub mIncluirUsuario()

    Dim strSql              As String
    Dim lngLinhasAfetadas   As Long
    Dim itmUsuarioInserido  As ListItem
    
    On Error GoTo ErrmIncluirUsuario
    
    Screen.MousePointer = vbHourglass
    
    If mValidarSenha Then
    
        If chkCLI.Value = 1 Then
            strSql = "SP_INCLUIR_USUARIO '" & txtLogin.Text & "', " & _
                     "'" & txtNomeUsuario.Text & "', " & _
                     "'" & txtSenha.Text & "', " & _
                     "'" & txtIntegrador.Text & "', " & _
                     Str(chkCLI.Value) & ", " & _
                     Trim(Str(cmbUnidade.ItemData(cmbUnidade.ListIndex))) & " "
        Else
            strSql = "SP_INCLUIR_USUARIO '" & txtLogin.Text & "', " & _
                     "'" & txtNomeUsuario.Text & "', " & _
                     "'" & txtSenha.Text & "', " & _
                     "'" & txtIntegrador.Text & "', " & _
                     Str(chkCLI.Value) & ", 0 "
        End If
        
        dbConexaoSMR.Execute strSql, lngLinhasAfetadas
        
        If lngLinhasAfetadas > 0 Then
            MsgBox "Usuário incluído com sucesso ! ", vbInformation + vbOKOnly, App.Title
        End If
        
        'Coloca o frame que contém a lista sobreposto ao frame de dados
        fraDadosCadAuxUsuarios.Visible = False
        fraListaCadAuxUsuarios.Visible = True
        
        'Adicionando o novo usuário na lista
        Set itmUsuarioInserido = lvwCadastroAuxUsuarios.ListItems.Add(, , txtLogin.Text)
        itmUsuarioInserido.SubItems(1) = txtNomeUsuario.Text
        itmUsuarioInserido.SubItems(2) = txtSenha.Text
        itmUsuarioInserido.SubItems(3) = txtIntegrador.Text
        
        If chkCLI.Value = 1 Then
            itmUsuarioInserido.SubItems(4) = True
            itmUsuarioInserido.SubItems(5) = cmbUnidade.ItemData(cmbUnidade.ListIndex)
        
        Else
            itmUsuarioInserido.SubItems(4) = False
            itmUsuarioInserido.SubItems(5) = 0
        
        End If
        
        itmUsuarioInserido.Selected = True
        itmUsuarioInserido.EnsureVisible
        
    Else
    
        MsgBox "Senha não confere ! ", vbInformation + vbOKOnly, App.Title
        txtSenha.SetFocus
    
    End If
    
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrmIncluirUsuario:

    Screen.MousePointer = vbNormal
    
    MsgBox "Ocorreu o seguinte erro na rotina mIncluirUsuario: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub

Private Sub mLimparDadosUsuario()

    txtLogin.Text = ""
    txtNomeUsuario.Text = ""
    txtSenha.Text = ""
    txtRedigiteSenha.Text = ""
    txtIntegrador.Text = ""
    chkCLI.Value = 0
    cmbUnidade.ListIndex = -1
    cmbUnidade.Enabled = False
    
End Sub


Private Sub mMontaCabecalhoListaUsuarios()

    'Criando os cabeçalhos das colunas da lista de usuários
    With lvwCadastroAuxUsuarios
        .ColumnHeaders.Add , , "Login", 1000
        .ColumnHeaders.Add , , "Nome", 7200
        .ColumnHeaders.Add , , "Senha", 0
        .ColumnHeaders.Add , , "Integrador", 0
        .ColumnHeaders.Add , , "CLI", 0
        .ColumnHeaders.Add , , "Unidade", 0
    End With
    
End Sub

Private Sub mPreencheListaUsuarios()

    Dim rsUsuarios  As ADODB.Recordset
    Dim itmUsuario  As ListItem
    Dim strSql      As String
    
    On Error GoTo ErrmPreencheListaUsuarios
    
    Screen.MousePointer = vbHourglass
    
    With lvwCadastroAuxUsuarios
    
        .ListItems.Clear
        .ColumnHeaders.Clear
        
        'Monta o cabeçalho da lista
        Call mMontaCabecalhoListaUsuarios
        
        strSql = "EXECUTE SP_LISTAR_USUARIO"
        
        Set rsUsuarios = New ADODB.Recordset
        rsUsuarios.Open strSql, dbConexaoSMR, adOpenStatic, adLockReadOnly, adCmdText
        
        Do While Not rsUsuarios.EOF
        
            Set itmUsuario = .ListItems.Add(, , rsUsuarios.Fields("Login").Value)
            itmUsuario.SubItems(1) = rsUsuarios.Fields("Nome_Usuario").Value
            itmUsuario.SubItems(2) = rsUsuarios.Fields("Senha").Value
            
            If IsNull(rsUsuarios.Fields("Integrador").Value) Then
                itmUsuario.SubItems(3) = ""
            Else
                itmUsuario.SubItems(3) = rsUsuarios.Fields("Integrador").Value
            End If
            
            itmUsuario.SubItems(4) = rsUsuarios.Fields("Flag_CLI").Value
            
            If IsNull(rsUsuarios.Fields("Cod_Unidade").Value) Then
                itmUsuario.SubItems(5) = 0
            Else
                itmUsuario.SubItems(5) = rsUsuarios.Fields("Cod_Unidade").Value
            End If
            
            rsUsuarios.MoveNext
            
        Loop
        
    End With
    
    rsUsuarios.Close
    Set rsUsuarios = Nothing
    
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrmPreencheListaUsuarios:
    
    Screen.MousePointer = vbNormal
    
    MsgBox "Ocorreu o seguinte erro na rotina mPreencheListaUsuarios: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title
           
    If rsUsuarios.State = adStateOpen Then
        rsUsuarios.Close
        Set rsUsuarios = Nothing
    End If

End Sub

Private Sub mTratarTelaAlterarUsuario()

    On Error GoTo ErrmTratarTelaAlterarUsuario
    
    'Desabilita os botões da ToolBar
    Call mDesabilitarBotoesToolBarUsuarios
        
    'Coloca o frame que contém os dados sobreposto ao frame de lista
    fraDadosCadAuxUsuarios.Visible = True
    fraListaCadAuxUsuarios.Visible = False
        
    Exit Sub
    
ErrmTratarTelaAlterarUsuario:

    MsgBox "Ocorreu o seguinte erro na rotina mTratarTelaAlterarUsuario: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub

Private Sub mTratarTelaExcluirUsuario()

    On Error GoTo ErrmTratarTelaExcluirUsuario
    
    'Coloca o frame que contém os dados sobreposto ao frame de lista
    fraDadosCadAuxUsuarios.Visible = True
    fraListaCadAuxUsuarios.Visible = False
    
    Exit Sub
    
ErrmTratarTelaExcluirUsuario:

    MsgBox "Ocorreu o seguinte erro na rotina mTratarTelaExcluirUsuario: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub


Private Sub mTratarTelaIncluirUsuario()

    On Error GoTo ErrmTratarTelaIncluirUsuario
    
    'Desabilita os botões da ToolBar
    Call mDesabilitarBotoesToolBarUsuarios
    
    'Limpa dados do usuário
    Call mLimparDadosUsuario
    
    'Coloca o frame que contém os dados sobreposto ao frame de lista
    fraDadosCadAuxUsuarios.Visible = True
    fraListaCadAuxUsuarios.Visible = False
    
    Exit Sub
    
ErrmTratarTelaIncluirUsuario:

    MsgBox "Ocorreu o seguinte erro na rotina mTratarTelaIncluirUsuario: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title
    
End Sub

Private Function mValidarCamposUsuario() As Boolean

    On Error GoTo ErrmValidarCamposUsuario
    
    If Trim(txtLogin.Text) = "" Then
        MsgBox "Login não foi preenchido !", vbInformation + vbOKOnly, App.Title
        txtLogin.SetFocus
        mValidarCamposUsuario = False
        Exit Function
    End If
    
    If Trim(txtNomeUsuario.Text) = "" Then
        MsgBox "Nome do usuário não foi preenchido !", vbInformation + vbOKOnly, App.Title
        txtNomeUsuario.SetFocus
        mValidarCamposUsuario = False
        Exit Function
    End If
    
    If Trim(txtSenha.Text) = "" Then
        MsgBox "Senha não foi preenchida !", vbInformation + vbOKOnly, App.Title
        txtSenha.SetFocus
        mValidarCamposUsuario = False
        Exit Function
    End If
    
    If chkCLI.Value = 1 Then
        If cmbUnidade.ListIndex = -1 Then
            MsgBox "Unidade não foi preenchida !", vbInformation + vbOKOnly, App.Title
            txtSenha.SetFocus
            mValidarCamposUsuario = False
            Exit Function
        End If
    End If
    
    mValidarCamposUsuario = True
        
    Exit Function
    
ErrmValidarCamposUsuario:

    MsgBox "Ocorreu o seguinte erro na rotina mValidarCamposUsuario: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title
    
End Function

Private Function mValidarSenha() As Boolean

    mValidarSenha = True
    
    If txtSenha.Text <> txtRedigiteSenha.Text Then
        mValidarSenha = False
    End If
    
End Function

Private Sub chkCLI_Click()

        If chkCLI.Value = 1 Then
            cmbUnidade.Enabled = True
        Else
            cmbUnidade.ListIndex = -1
            cmbUnidade.Enabled = False
        End If

End Sub

Private Sub cmdCancelar_Click()

    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadAuxUsuarios.Visible = False
    fraListaCadAuxUsuarios.Visible = True

    'Habilita os botões da ToolBar
    Call mHabilitarBotoesToolBarUsuarios

End Sub

Private Sub cmdGravar_Click()
    
    On Error GoTo ErrcmdGravar_Click
    
    If mValidarCamposUsuario() Then
    
        If fraDadosCadAuxUsuarios.Tag = "I" Then
        
            'Inclusão
            Call mIncluirUsuario
        
        ElseIf fraDadosCadAuxUsuarios.Tag = "A" Then
            
            'Alteração
            Call mAlterarUsuario
        
        End If
        
        'Habilita os botões da ToolBar
        Call mHabilitarBotoesToolBarUsuarios
        
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
    Call mHabilitarBotoesToolBarUsuarios

    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadAuxUsuarios.Visible = False
    fraListaCadAuxUsuarios.Visible = True

    'Preenche a combo de unidade
    Call mPreencherComboUnidade
    
    'Preenche a lista de usuários
    Call mPreencheListaUsuarios
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmPrincipal.StatusBar.Panels(1).Text = ""
    
End Sub


Private Sub lvwCadastroAuxUsuarios_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    With lvwCadastroAuxUsuarios
    
        .SortKey = ColumnHeader.Index - 1
        
        If .SortOrder = lvwAscending Then
           .SortOrder = lvwDescending
        Else
           .SortOrder = lvwAscending
        End If
        
        .Sorted = True
        
    End With

End Sub


Private Sub tlbCadastroAuxUsuarios_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
    
        Case 1 'Incluir
        
            Call mTratarTelaIncluirUsuario
            Call mHabilitarDesabilitarCamposCadastroUsuario("I")
                    
        Case 2 'Alterar
        
            Call mTratarTelaAlterarUsuario
            Call mDadosUsuario
            Call mHabilitarDesabilitarCamposCadastroUsuario("A")
                        
        Case 3 'Excluir
        
            If MsgBox("Deseja realmente excluir este registro ?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                Call mExcluirUsuario
            End If
            
        Case 5 'Sair
            
            Unload Me
            
    End Select

End Sub

Private Sub mPreencherComboUnidade()

    Dim rsUnidade  As ADODB.Recordset
    Dim strSql      As String

    On Error GoTo ErrmPreencherComboUnidade
    
    strSql = "EXECUTE SP_LISTAR_UNIDADE"
    
    Set rsUnidade = New ADODB.Recordset
    rsUnidade.Open strSql, dbConexaoSMR, adOpenStatic, adLockReadOnly, adCmdText
    
    cmbUnidade.AddItem ""
    cmbUnidade.ItemData(cmbUnidade.NewIndex) = 0
    
    Do While Not rsUnidade.EOF
    
        cmbUnidade.AddItem rsUnidade.Fields("Desc_Unidade").Value
        cmbUnidade.ItemData(cmbUnidade.NewIndex) = rsUnidade.Fields("Cod_Unidade").Value
        rsUnidade.MoveNext
        
    Loop
    
    cmbUnidade.ListIndex = 0
    
    rsUnidade.Close
    Set rsUnidade = Nothing
    
    Exit Sub
    
ErrmPreencherComboUnidade:
    
    MsgBox "Ocorreu o seguinte erro na rotina mPreencherComboUnidade: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title
           
    If rsUnidade.State = adStateOpen Then
        rsUnidade.Close
        Set rsUnidade = Nothing
    End If
    
End Sub


