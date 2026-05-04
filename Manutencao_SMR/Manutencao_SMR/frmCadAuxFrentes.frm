VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCadAuxFrentes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Frentes"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10620
   Icon            =   "frmCadAuxFrentes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.Toolbar tlbCadastroAuxFrentes 
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
   Begin VB.Frame fraListaCadAuxFrentes 
      Caption         =   "Lista de frentes cadastradas"
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
               Picture         =   "frmCadAuxFrentes.frx":0442
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAuxFrentes.frx":0554
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAuxFrentes.frx":0666
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAuxFrentes.frx":07D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAuxFrentes.frx":08E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAuxFrentes.frx":0D34
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvwCadastroAuxFrentes 
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
   Begin VB.Frame fraDadosCadAuxFrentes 
      Caption         =   "Dados da frente"
      Height          =   2835
      Left            =   15
      TabIndex        =   4
      Top             =   540
      Width           =   10575
      Begin VB.Frame Frame1 
         Height          =   915
         Left            =   105
         TabIndex        =   6
         Top             =   1785
         Width           =   10365
         Begin VB.CommandButton cmdGravar 
            Caption         =   "Gravar"
            Height          =   375
            Left            =   7950
            TabIndex        =   8
            Top             =   315
            Width           =   1020
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   9180
            TabIndex        =   7
            Top             =   315
            Width           =   1020
         End
      End
      Begin VB.TextBox txtDescricaoFrente 
         Height          =   345
         Left            =   2565
         MaxLength       =   50
         TabIndex        =   1
         Top             =   495
         Width           =   6675
      End
      Begin VB.Label lblDescricaoFrente 
         AutoSize        =   -1  'True
         Caption         =   "Descrição da Frente:"
         Height          =   195
         Left            =   960
         TabIndex        =   5
         Top             =   570
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmCadAuxFrentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub mDesabilitarBotoesToolBarFrentes()

    With tlbCadastroAuxFrentes
    
        .Buttons(1).Enabled = False
        .Buttons(2).Enabled = False
        .Buttons(3).Enabled = False
        
    End With
    
End Sub

Private Sub mHabilitarBotoesToolBarFrentes()

    With tlbCadastroAuxFrentes
    
        .Buttons(1).Enabled = True
        .Buttons(2).Enabled = True
        .Buttons(3).Enabled = True
        
    End With
    
End Sub

Private Sub mAlterarFrente()

    Dim strSql            As String
    Dim lngLinhasAfetadas As Long
    
    On Error GoTo ErrmAlterarFrente
    
    Screen.MousePointer = vbHourglass
    
    strSql = "SP_ALTERAR_FRENTE " & CInt(lvwCadastroAuxFrentes.SelectedItem.Text) & ", " & _
             "'" & txtDescricaoFrente.Text & "'"
             
    dbConexaoSMR.Execute strSql, lngLinhasAfetadas
    
    If lngLinhasAfetadas > 0 Then
        MsgBox "Frente alterada com sucesso ! ", vbInformation + vbOKOnly, App.Title
    End If
    
    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadAuxFrentes.Visible = False
    fraListaCadAuxFrentes.Visible = True
    
    'Alterando os dados do usuário na lista
    With lvwCadastroAuxFrentes.SelectedItem
        .SubItems(1) = txtDescricaoFrente.Text
        .EnsureVisible
    End With
        
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrmAlterarFrente:

    Screen.MousePointer = vbNormal
    
    MsgBox "Ocorreu o seguinte erro na rotina mAlterarFrente: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub

Private Sub mDadosFrente()

    On Error GoTo ErrmDadosFrente
    
    With lvwCadastroAuxFrentes.SelectedItem
        txtDescricaoFrente.Text = .SubItems(1)
    End With
    
    Exit Sub
    
ErrmDadosFrente:

    MsgBox "Ocorreu o seguinte erro na rotina mDadosFrente: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub



Private Sub mExcluirFrente()

    Dim strSql            As String
    Dim lngLinhasAfetadas As Long
    
    On Error GoTo ErrmExcluirFrente
    
    Screen.MousePointer = vbHourglass
    
    strSql = "SP_EXCLUIR_FRENTE " & CInt(lvwCadastroAuxFrentes.SelectedItem.Text)
             
    dbConexaoSMR.Execute strSql, lngLinhasAfetadas
    
    If lngLinhasAfetadas > 0 Then
        MsgBox "Frente excluída com sucesso ! ", vbInformation + vbOKOnly, App.Title
    End If
    
    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadAuxFrentes.Visible = False
    fraListaCadAuxFrentes.Visible = True
    
    'Preencher lista de frentes
    Call mPreencheListaFrentes
    
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrmExcluirFrente:

    Screen.MousePointer = vbNormal
    
    MsgBox "Ocorreu o seguinte erro na rotina mExcluirFrente: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub

Private Sub mHabilitarDesabilitarCamposCadastroFrente(ByVal pAcao As String)

    fraDadosCadAuxFrentes.Tag = pAcao

    If UCase(Trim(pAcao)) = "I" Then
        
        txtDescricaoFrente.Enabled = True
        
        txtDescricaoFrente.SetFocus
    
    ElseIf UCase(Trim(pAcao)) = "A" Then
        
        txtDescricaoFrente.Enabled = True
        
    End If
    
End Sub

Private Sub mIncluirFrente()

    Dim strSql              As String
    Dim lngCodigoFrente     As Long
    Dim lngLinhasAfetadas   As Long
    Dim itmFrenteInserida   As ListItem
    Dim cmdFrente           As ADODB.Command
    
    On Error GoTo ErrmIncluirFrente
    
    Screen.MousePointer = vbHourglass
    
    Set cmdFrente = New ADODB.Command
    
    With cmdFrente
    
        .ActiveConnection = dbConexaoSMR
        .CommandType = adCmdStoredProc
        .CommandText = "SP_INCLUIR_FRENTE"
        
        .Parameters.Refresh
        .Parameters(1).Value = Trim(txtDescricaoFrente.Text)
        
        .Execute
        
        lngCodigoFrente = .Parameters(2).Value
        
    End With
        
    If lngCodigoFrente > 0 Then
        MsgBox "Frente incluída com sucesso ! ", vbInformation + vbOKOnly, App.Title
    End If
    
    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadAuxFrentes.Visible = False
    fraListaCadAuxFrentes.Visible = True
    
    'Adicionando a nova página na lista
    Set itmFrenteInserida = lvwCadastroAuxFrentes.ListItems.Add(, , Format(lngCodigoFrente, "000"))
    itmFrenteInserida.SubItems(1) = txtDescricaoFrente.Text
    
    itmFrenteInserida.Selected = True
    itmFrenteInserida.EnsureVisible
        
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrmIncluirFrente:

    Screen.MousePointer = vbNormal
    
    MsgBox "Ocorreu o seguinte erro na rotina mIncluirFrente: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub

Private Sub mLimparDadosFrente()

    txtDescricaoFrente.Text = ""
    
End Sub


Private Sub mMontaCabecalhoListaFrentes()

    'Criando os cabeçalhos das colunas da lista de frentes
    With lvwCadastroAuxFrentes
        .ColumnHeaders.Add , , "Código", 1000
        .ColumnHeaders.Add , , "Descrição", 9370
    End With
    
End Sub

Private Sub mPreencheListaFrentes()

    Dim rsFrentes    As ADODB.Recordset
    Dim itmFrente    As ListItem
    Dim strSql       As String
    
    On Error GoTo ErrmPreencheListaFrentes
    
    Screen.MousePointer = vbHourglass
    
    With lvwCadastroAuxFrentes
    
        .ListItems.Clear
        .ColumnHeaders.Clear
        
        'Monta o cabeçalho da lista
        Call mMontaCabecalhoListaFrentes
        
        strSql = "EXECUTE SP_LISTAR_FRENTES"
        
        Set rsFrentes = New ADODB.Recordset
        rsFrentes.Open strSql, dbConexaoSMR, adOpenStatic, adLockReadOnly, adCmdText
        
        Do While Not rsFrentes.EOF
        
            Set itmFrente = .ListItems.Add(, , Format(rsFrentes.Fields("Cod_Frente").Value, "000"))
            itmFrente.SubItems(1) = rsFrentes.Fields("Desc_Frente").Value
              
            rsFrentes.MoveNext
            
        Loop
        
    End With
    
    rsFrentes.Close
    Set rsFrentes = Nothing
    
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrmPreencheListaFrentes:

    Screen.MousePointer = vbNormal
    
    MsgBox "Ocorreu o seguinte erro na rotina mPreencheListaFrentes: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title
           
    If rsFrentes.State = adStateOpen Then
        rsFrentes.Close
        Set rsFrentes = Nothing
    End If

End Sub

Private Sub mTratarTelaAlterarFrente()

    On Error GoTo ErrmTratarTelaAlterarFrente
    
    'Desabilita os botões da ToolBar
    Call mDesabilitarBotoesToolBarFrentes
    
    'Coloca o frame que contém os dados sobreposto ao frame de lista
    fraDadosCadAuxFrentes.Visible = True
    fraListaCadAuxFrentes.Visible = False
        
    Exit Sub
    
ErrmTratarTelaAlterarFrente:

    MsgBox "Ocorreu o seguinte erro na rotina mTratarTelaAlterarFrente: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub

Private Sub mTratarTelaExcluirFrente()

    On Error GoTo ErrmTratarTelaExcluirFrente
    
    'Coloca o frame que contém os dados sobreposto ao frame de lista
    fraDadosCadAuxFrentes.Visible = True
    fraListaCadAuxFrentes.Visible = False
    
    Exit Sub
    
ErrmTratarTelaExcluirFrente:

    MsgBox "Ocorreu o seguinte erro na rotina mTratarTelaExcluirFrente: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub


Private Sub mTratarTelaIncluirFrente()

    On Error GoTo ErrmTratarTelaIncluirFrente
    
    'Desabilita os botões da ToolBar
    Call mDesabilitarBotoesToolBarFrentes
    
    'Limpa dados da frente
    Call mLimparDadosFrente
    
    'Coloca o frame que contém os dados sobreposto ao frame de lista
    fraDadosCadAuxFrentes.Visible = True
    fraListaCadAuxFrentes.Visible = False
    
    Exit Sub
    
ErrmTratarTelaIncluirFrente:

    MsgBox "Ocorreu o seguinte erro na rotina mTratarTelaIncluirFrente: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title
    
End Sub

Private Function mValidarCamposFrente() As Boolean

    On Error GoTo ErrmValidarCamposFrente
    
    If Trim(txtDescricaoFrente.Text) = "" Then
        MsgBox "Descrição da frente não foi preenchida !", vbInformation + vbOKOnly, App.Title
        txtDescricaoFrente.SetFocus
        mValidarCamposFrente = False
        Exit Function
    End If
        
    mValidarCamposFrente = True
        
    Exit Function
    
ErrmValidarCamposFrente:

    MsgBox "Ocorreu o seguinte erro na rotina mValidarCamposFrente: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title
    
End Function

Private Sub cmdCancelar_Click()

    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadAuxFrentes.Visible = False
    fraListaCadAuxFrentes.Visible = True

    'Habilita os botões da ToolBar
    Call mHabilitarBotoesToolBarFrentes

End Sub

Private Sub cmdGravar_Click()
    
    On Error GoTo ErrcmdGravar_Click
    
    If mValidarCamposFrente() Then
    
        If fraDadosCadAuxFrentes.Tag = "I" Then
        
            'Inclusao
            Call mIncluirFrente
        
        ElseIf fraDadosCadAuxFrentes.Tag = "A" Then
            
            'Alteração
            Call mAlterarFrente
        
        End If
        
        'Habilita os botões da ToolBar
        Call mHabilitarBotoesToolBarFrentes
        
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
    Call mHabilitarBotoesToolBarFrentes
        
    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadAuxFrentes.Visible = False
    fraListaCadAuxFrentes.Visible = True

    'Preenche a lista de Frentes
    Call mPreencheListaFrentes
    
End Sub






Private Sub Form_Unload(Cancel As Integer)

    frmPrincipal.StatusBar.Panels(1).Text = ""
    
End Sub


Private Sub lvwCadastroAuxFrentes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    With lvwCadastroAuxFrentes
    
        .SortKey = ColumnHeader.Index - 1
        
        If .SortOrder = lvwAscending Then
           .SortOrder = lvwDescending
        Else
           .SortOrder = lvwAscending
        End If
        
        .Sorted = True
        
    End With

End Sub


Private Sub tlbCadastroAuxFrentes_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
    
        Case 1 'Incluir
        
            Call mTratarTelaIncluirFrente
            Call mHabilitarDesabilitarCamposCadastroFrente("I")
                    
        Case 2 'Alterar
        
            Call mTratarTelaAlterarFrente
            Call mDadosFrente
            Call mHabilitarDesabilitarCamposCadastroFrente("A")
                        
        Case 3 'Excluir
        
            If MsgBox("Deseja realmente excluir este registro ?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                Call mExcluirFrente
            End If
            
        Case 5 'Sair
            
            Unload Me
            
    End Select

End Sub


