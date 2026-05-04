VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCadProjetosPMO 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Projetos do PMO"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10785
   Icon            =   "frmCadProjetosPMO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraListaCadProjetosPMO 
      Caption         =   "Lista de Projetos PMO cadastrados"
      Height          =   2880
      Left            =   15
      TabIndex        =   4
      Top             =   510
      Width           =   10740
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   9585
         Top             =   1650
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
               Picture         =   "frmCadProjetosPMO.frx":0442
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadProjetosPMO.frx":0554
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadProjetosPMO.frx":0666
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadProjetosPMO.frx":07D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadProjetosPMO.frx":08E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadProjetosPMO.frx":0D34
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvwCadastroProjetosPMO 
         Height          =   2490
         Left            =   45
         TabIndex        =   5
         Top             =   255
         Width           =   10620
         _ExtentX        =   18733
         _ExtentY        =   4392
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
   Begin MSComctlLib.Toolbar tlbCadastroProjetosPMO 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10785
      _ExtentX        =   19024
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
   Begin VB.Frame fraDadosCadProjetosPMO 
      Caption         =   "Dados do Projeto PMO"
      Height          =   2865
      Left            =   15
      TabIndex        =   6
      Top             =   525
      Width           =   10740
      Begin VB.Frame Frame1 
         Height          =   885
         Left            =   120
         TabIndex        =   10
         Top             =   1785
         Width           =   10500
         Begin VB.CommandButton cmdGravar 
            Caption         =   "Gravar"
            Height          =   375
            Left            =   8055
            TabIndex        =   12
            Top             =   315
            Width           =   1020
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   9285
            TabIndex        =   11
            Top             =   315
            Width           =   1020
         End
      End
      Begin VB.TextBox txtNomeProjetoPMO 
         Height          =   345
         Left            =   2445
         MaxLength       =   50
         TabIndex        =   2
         Top             =   900
         Width           =   6675
      End
      Begin VB.TextBox txtOrdenacao 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   2445
         MaxLength       =   2
         TabIndex        =   3
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txtCodigoProjetoPMO 
         Height          =   345
         Left            =   2445
         MaxLength       =   25
         TabIndex        =   1
         Top             =   480
         Width           =   4395
      End
      Begin VB.Label lblOrdenacao 
         AutoSize        =   -1  'True
         Caption         =   "Ordenação:"
         Height          =   195
         Left            =   1515
         TabIndex        =   9
         Top             =   1395
         Width           =   840
      End
      Begin VB.Label lblCodigoProjetoPMO 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   1845
         TabIndex        =   8
         Top             =   540
         Width           =   540
      End
      Begin VB.Label lblNomeProjetoPMO 
         AutoSize        =   -1  'True
         Caption         =   "Nome do Projeto:"
         Height          =   195
         Left            =   1095
         TabIndex        =   7
         Top             =   975
         Width           =   1230
      End
   End
End
Attribute VB_Name = "frmCadProjetosPMO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mHabilitarBotoesToolBarProjetosPMO()

    With tlbCadastroProjetosPMO
    
        .Buttons(1).Enabled = True
        .Buttons(2).Enabled = True
        .Buttons(3).Enabled = True
        
    End With
    
End Sub

Private Sub mDesabilitarBotoesToolBarProjetosPMO()

    With tlbCadastroProjetosPMO
    
        .Buttons(1).Enabled = False
        .Buttons(2).Enabled = False
        .Buttons(3).Enabled = False
        
    End With
    
End Sub

Private Function mValidarCamposProjetoPMO() As Boolean

    On Error GoTo ErrmValidarCamposProjetoPMO
    
    If Trim(txtCodigoProjetoPMO.Text) = "" Then
        MsgBox "Código do projeto PMO não foi preenchido !", vbInformation + vbOKOnly, App.Title
        txtCodigoProjetoPMO.SetFocus
        mValidarCamposProjetoPMO = False
        Exit Function
    End If
    
    If Trim(txtNomeProjetoPMO.Text) = "" Then
        MsgBox "Nome do projeto PMO não foi preenchido !", vbInformation + vbOKOnly, App.Title
        txtNomeProjetoPMO.SetFocus
        mValidarCamposProjetoPMO = False
        Exit Function
    End If
    
    If Trim(txtOrdenacao.Text) = "" Then
        MsgBox "Ordenação não foi preenchida !", vbInformation + vbOKOnly, App.Title
        txtOrdenacao.SetFocus
        mValidarCamposProjetoPMO = False
        Exit Function
    End If
    
    mValidarCamposProjetoPMO = True
        
    Exit Function
    
ErrmValidarCamposProjetoPMO:

    MsgBox "Ocorreu o seguinte erro na rotina mValidarCamposProjetoPMO: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title
    
End Function

Private Sub mAlterarProjetoPMO()

    Dim strSql                      As String
    Dim lngLinhasAfetadas           As Long
    Dim rsOrdenacaoProjetoExiste    As ADODB.Recordset
    
    On Error GoTo ErrmAlterarProjetoPMO
    
    Screen.MousePointer = vbHourglass
    
    strSql = "EXECUTE SP_VERIFICAR_ORDENACAO_PROJETO_EXISTE " & CInt(txtOrdenacao.Text)
    
    Set rsOrdenacaoProjetoExiste = New ADODB.Recordset
    rsOrdenacaoProjetoExiste.Open strSql, dbConexaoSMR, adOpenStatic, adLockReadOnly, adCmdText
    
    If UCase(Trim(rsOrdenacaoProjetoExiste.Fields("Project_Code").Value)) = UCase(Trim(txtCodigoProjetoPMO.Text)) Then
    
        strSql = "SP_ALTERAR_PROJETO '" & txtCodigoProjetoPMO.Text & "', " & _
                 "'" & txtNomeProjetoPMO.Text & "', " & _
                 CInt(txtOrdenacao.Text)
                 
        dbConexaoSMR.Execute strSql, lngLinhasAfetadas
        
        If lngLinhasAfetadas > 0 Then
            MsgBox "Projeto PMO alterado com sucesso ! ", vbInformation + vbOKOnly, App.Title
        End If
        
        'Coloca o frame que contém a lista sobreposto ao frame de dados
        fraDadosCadProjetosPMO.Visible = False
        fraListaCadProjetosPMO.Visible = True
        
        'Alterando os dados do usuário na lista
        With lvwCadastroProjetosPMO.SelectedItem
            .Text = txtCodigoProjetoPMO.Text
            .SubItems(1) = txtNomeProjetoPMO.Text
            .SubItems(2) = Format(txtOrdenacao.Text, "00")
            .EnsureVisible
        End With
        
    Else
    
        MsgBox "Já existe um outro Projeto PMO com essa ordenação ! ", vbInformation + vbOKOnly, App.Title
        txtOrdenacao.SetFocus
    
    End If
    
    rsOrdenacaoProjetoExiste.Close
    Set rsOrdenacaoProjetoExiste = Nothing
        
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrmAlterarProjetoPMO:

    Screen.MousePointer = vbNormal
    
    MsgBox "Ocorreu o seguinte erro na rotina mAlterarProjetoPMO: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title
           
    If rsOrdenacaoProjetoExiste.State = adStateOpen Then
        rsOrdenacaoProjetoExiste.Close
        Set rsOrdenacaoProjetoExiste = Nothing
    End If

End Sub

Private Sub mIncluirProjetoPMO()

    Dim strSql                      As String
    Dim lngLinhasAfetadas           As Long
    Dim itmProjetoPMOInserido       As ListItem
    Dim rsOrdenacaoProjetoExiste    As ADODB.Recordset
    
    On Error GoTo ErrmIncluirProjetoPMO
    
    Screen.MousePointer = vbHourglass
    
    strSql = "EXECUTE SP_VERIFICAR_ORDENACAO_PROJETO_EXISTE " & CInt(txtOrdenacao.Text)
    
    Set rsOrdenacaoProjetoExiste = New ADODB.Recordset
    rsOrdenacaoProjetoExiste.Open strSql, dbConexaoSMR, adOpenStatic, adLockReadOnly, adCmdText
    
    If rsOrdenacaoProjetoExiste.RecordCount = 0 Then
    
        strSql = "SP_INCLUIR_PROJETO '" & txtCodigoProjetoPMO.Text & "', " & _
                 "'" & txtNomeProjetoPMO.Text & "', " & _
                 CInt(txtOrdenacao.Text)
                 
        dbConexaoSMR.Execute strSql, lngLinhasAfetadas
        
        If lngLinhasAfetadas > 0 Then
            MsgBox "Projeto PMO incluído com sucesso ! ", vbInformation + vbOKOnly, App.Title
        End If
        
        'Coloca o frame que contém a lista sobreposto ao frame de dados
        fraDadosCadProjetosPMO.Visible = False
        fraListaCadProjetosPMO.Visible = True
        
        'Adicionando o novo usuário na lista
        Set itmProjetoPMOInserido = lvwCadastroProjetosPMO.ListItems.Add(, , txtCodigoProjetoPMO.Text)
        itmProjetoPMOInserido.SubItems(1) = txtNomeProjetoPMO.Text
        itmProjetoPMOInserido.SubItems(2) = Format(txtOrdenacao.Text, "00")
        
        itmProjetoPMOInserido.Selected = True
        itmProjetoPMOInserido.EnsureVisible
        
    Else
        
        MsgBox "Já existe um Projeto PMO com essa ordenação ! ", vbInformation + vbOKOnly, App.Title
        txtOrdenacao.SetFocus
    
    End If
    
        
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrmIncluirProjetoPMO:

    Screen.MousePointer = vbNormal
    
    MsgBox "Ocorreu o seguinte erro na rotina mIncluirProjetoPMO: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub

Private Sub mExcluirProjetoPMO()

    Dim strSql                      As String
    Dim lngLinhasAfetadas           As Long
    
    On Error GoTo ErrmExcluirProjetoPMO
    
    Screen.MousePointer = vbHourglass
        
    strSql = "SP_EXCLUIR_PROJETO '" & lvwCadastroProjetosPMO.SelectedItem.Text & "'"
             
    dbConexaoSMR.Execute strSql, lngLinhasAfetadas
    
    If lngLinhasAfetadas > 0 Then
        MsgBox "Projeto PMO excluído com sucesso ! ", vbInformation + vbOKOnly, App.Title
    End If
    
    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadProjetosPMO.Visible = False
    fraListaCadProjetosPMO.Visible = True
    
    'Preencher lista de usuários
    Call mPreencheListaProjetosPMO
    
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrmExcluirProjetoPMO:

    Screen.MousePointer = vbNormal
    
    MsgBox "Ocorreu o seguinte erro na rotina mExcluirProjetoPMO: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub


Private Sub mDadosProjetoPMO()

    On Error GoTo ErrmDadosProjetoPMO
    
    With lvwCadastroProjetosPMO.SelectedItem
        txtCodigoProjetoPMO.Text = .Text
        txtNomeProjetoPMO.Text = .SubItems(1)
        txtOrdenacao.Text = .SubItems(2)
    End With
    
    Exit Sub
    
ErrmDadosProjetoPMO:

    MsgBox "Ocorreu o seguinte erro na rotina mDadosProjetoPMO: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub

Private Sub mTratarTelaIncluirProjetoPMO()

    On Error GoTo ErrmTratarTelaIncluirProjetoPMO
    
    'Desabilita os botões da ToolBar
    Call mDesabilitarBotoesToolBarProjetosPMO
    
    'Limpa dados
    Call mLimparDadosProjetoPMO
    
    'Coloca o frame que contém os dados sobreposto ao frame de lista
    fraDadosCadProjetosPMO.Visible = True
    fraListaCadProjetosPMO.Visible = False
    
    Exit Sub
    
ErrmTratarTelaIncluirProjetoPMO:

    MsgBox "Ocorreu o seguinte erro na rotina mTratarTelaIncluirProjetoPMO: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title
    
End Sub

Private Sub mLimparDadosProjetoPMO()

    txtCodigoProjetoPMO.Text = ""
    txtNomeProjetoPMO.Text = ""
    txtOrdenacao.Text = ""
    
End Sub

Private Sub mTratarTelaAlterarProjetoPMO()

    On Error GoTo ErrmTratarTelaAlterarProjetoPMO
    
    'Desabilita os botões da ToolBar
    Call mDesabilitarBotoesToolBarProjetosPMO
    
    'Coloca o frame que contém os dados sobreposto ao frame de lista
    fraDadosCadProjetosPMO.Visible = True
    fraListaCadProjetosPMO.Visible = False
        
    Exit Sub
    
ErrmTratarTelaAlterarProjetoPMO:

    MsgBox "Ocorreu o seguinte erro na rotina mTratarTelaAlterarProjetoPMO: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub

Private Sub mTratarTelaExcluirProjetoPMO()

    On Error GoTo ErrmTratarTelaExcluirProjetoPMO
    
    'Coloca o frame que contém os dados sobreposto ao frame de lista
    fraDadosCadProjetosPMO.Visible = True
    fraListaCadProjetosPMO.Visible = False
    
    Exit Sub
    
ErrmTratarTelaExcluirProjetoPMO:

    MsgBox "Ocorreu o seguinte erro na rotina mTratarTelaExcluirProjetoPMO: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub

Private Sub mHabilitarDesabilitarCamposCadastroProjetoPMO(ByVal pAcao As String)

    fraDadosCadProjetosPMO.Tag = pAcao

    If UCase(Trim(pAcao)) = "I" Then
        
        txtCodigoProjetoPMO.Enabled = True
        txtNomeProjetoPMO.Enabled = True
        txtOrdenacao.Enabled = True
        
        txtCodigoProjetoPMO.SetFocus
    
    ElseIf UCase(Trim(pAcao)) = "A" Then
        
        txtCodigoProjetoPMO.Enabled = False
        txtNomeProjetoPMO.Enabled = True
        txtOrdenacao.Enabled = True
        
    End If
    
End Sub

Private Sub mPreencheListaProjetosPMO()

    Dim rsProjetosPMO   As ADODB.Recordset
    Dim itmProjetoPMO   As ListItem
    Dim strSql          As String
    
    On Error GoTo ErrmPreencheListaProjetosPMO
    
    Screen.MousePointer = vbHourglass
    
    With lvwCadastroProjetosPMO
    
        .ListItems.Clear
        .ColumnHeaders.Clear
    
        'Monta o cabeçalho da lista
        Call mMontaCabecalhoListaProjetosPMO
        
        strSql = "EXECUTE SP_LISTAR_PROJETOS"
        
        Set rsProjetosPMO = New ADODB.Recordset
        rsProjetosPMO.Open strSql, dbConexaoSMR, adOpenStatic, adLockReadOnly, adCmdText
        
        Do While Not rsProjetosPMO.EOF
        
            Set itmProjetoPMO = .ListItems.Add(, , rsProjetosPMO.Fields("Project_Code").Value)
            itmProjetoPMO.SubItems(1) = rsProjetosPMO.Fields("Desc_Projeto").Value
            itmProjetoPMO.SubItems(2) = Format(rsProjetosPMO.Fields("Prioridade_Ordenacao").Value, "00")
              
            rsProjetosPMO.MoveNext
            
        Loop
    
    End With
    
    rsProjetosPMO.Close
    Set rsProjetosPMO = Nothing
    
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrmPreencheListaProjetosPMO:

    Screen.MousePointer = vbNormal
    
    MsgBox "Ocorreu o seguinte erro na rotina mPreencheListaProjetosPMO: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title
           
    If rsProjetosPMO.State = adStateOpen Then
        rsProjetosPMO.Close
        Set rsProjetosPMO = Nothing
    End If

End Sub

Private Sub mMontaCabecalhoListaProjetosPMO()

    'Criando os cabeçalhos das colunas da lista de usuários
    With lvwCadastroProjetosPMO
        .ColumnHeaders.Add , , "Código", 2000
        .ColumnHeaders.Add , , "Nome", 7200
        .ColumnHeaders.Add , , "Ordenação", 1100
    End With
    
End Sub

Private Sub cmdCancelar_Click()

    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadProjetosPMO.Visible = False
    fraListaCadProjetosPMO.Visible = True

    'Habilita os botões da ToolBar
    Call mHabilitarBotoesToolBarProjetosPMO

End Sub

Private Sub cmdGravar_Click()

    On Error GoTo ErrcmdGravar_Click
    
    If mValidarCamposProjetoPMO() Then
    
        If fraDadosCadProjetosPMO.Tag = "I" Then
        
            'Inclusao
            Call mIncluirProjetoPMO
        
        ElseIf fraDadosCadProjetosPMO.Tag = "A" Then
            
            'Alteração
            Call mAlterarProjetoPMO
                
        End If
        
        'Habilita os botões da ToolBar
        Call mHabilitarBotoesToolBarProjetosPMO
        
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
    Call mHabilitarBotoesToolBarProjetosPMO

    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadProjetosPMO.Visible = False
    fraListaCadProjetosPMO.Visible = True

    'Preenche a lista de Projetos PMO
    Call mPreencheListaProjetosPMO

End Sub


Private Sub Form_Unload(Cancel As Integer)

    frmPrincipal.StatusBar.Panels(1).Text = ""
    
End Sub


Private Sub lvwCadastroProjetosPMO_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    With lvwCadastroProjetosPMO
    
        .SortKey = ColumnHeader.Index - 1
        
        If .SortOrder = lvwAscending Then
           .SortOrder = lvwDescending
        Else
           .SortOrder = lvwAscending
        End If
        
        .Sorted = True
        
    End With

End Sub


Private Sub tlbCadastroProjetosPMO_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
    
        Case 1 'Incluir
        
            Call mTratarTelaIncluirProjetoPMO
            Call mHabilitarDesabilitarCamposCadastroProjetoPMO("I")
                    
        Case 2 'Alterar
        
            Call mTratarTelaAlterarProjetoPMO
            Call mDadosProjetoPMO
            Call mHabilitarDesabilitarCamposCadastroProjetoPMO("A")
                        
        Case 3 'Excluir
        
            If MsgBox("Deseja realmente excluir este registro ?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                Call mExcluirProjetoPMO
            End If
            
        Case 5 'Sair
            
            Unload Me
            
    End Select

End Sub


