VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCadAuxArea 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Areas"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10620
   Icon            =   "frmCadAuxArea.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.Toolbar tlbCadastroAuxArea 
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
   Begin VB.Frame fraListaCadAuxArea 
      Caption         =   "Lista de areas cadastradas"
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
               Picture         =   "frmCadAuxArea.frx":0442
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAuxArea.frx":0554
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAuxArea.frx":0666
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAuxArea.frx":07D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAuxArea.frx":08E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAuxArea.frx":0D34
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvwCadastroAuxArea 
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
   Begin VB.Frame fraDadosCadAuxArea 
      Caption         =   "Dados da area"
      Height          =   2835
      Left            =   15
      TabIndex        =   4
      Top             =   540
      Width           =   10575
      Begin VB.Frame Frame1 
         Height          =   795
         Left            =   120
         TabIndex        =   6
         Top             =   1905
         Width           =   10350
         Begin VB.CommandButton cmdGravar 
            Caption         =   "Gravar"
            Height          =   375
            Left            =   7845
            TabIndex        =   8
            Top             =   255
            Width           =   1020
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   9075
            TabIndex        =   7
            Top             =   255
            Width           =   1020
         End
      End
      Begin VB.TextBox txtDescricaoArea 
         Height          =   345
         Left            =   2925
         MaxLength       =   50
         TabIndex        =   1
         Top             =   495
         Width           =   6675
      End
      Begin VB.Label lblDescricaoArea 
         AutoSize        =   -1  'True
         Caption         =   "Descrição da Area:"
         Height          =   195
         Left            =   945
         TabIndex        =   5
         Top             =   570
         Width           =   1365
      End
   End
End
Attribute VB_Name = "frmCadAuxArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub mDesabilitarBotoesToolBarArea()

    With tlbCadastroAuxArea
    
        .Buttons(1).Enabled = False
        .Buttons(2).Enabled = False
        .Buttons(3).Enabled = False
        
    End With
    
End Sub

Private Sub mHabilitarBotoesToolBarArea()

    With tlbCadastroAuxArea
    
        .Buttons(1).Enabled = True
        .Buttons(2).Enabled = True
        .Buttons(3).Enabled = True
        
    End With
    
End Sub

Private Sub mAlterarArea()

    Dim strSql            As String
    Dim lngLinhasAfetadas As Long
    
    On Error GoTo ErrmAlterarArea
    
    Screen.MousePointer = vbHourglass
    
    strSql = "SP_ALTERAR_Area " & CInt(lvwCadastroAuxArea.SelectedItem.Text) & ", " & _
             "'" & txtDescricaoArea.Text & "'"
             
    dbConexaoSMR.Execute strSql, lngLinhasAfetadas
    
    If lngLinhasAfetadas > 0 Then
        MsgBox "Area alterada com sucesso ! ", vbInformation + vbOKOnly, App.Title
    End If
    
    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadAuxArea.Visible = False
    fraListaCadAuxArea.Visible = True
    
    'Alterando os dados na lista
    With lvwCadastroAuxArea.SelectedItem
        .SubItems(1) = txtDescricaoArea.Text
        .EnsureVisible
    End With
        
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrmAlterarArea:

    Screen.MousePointer = vbNormal
    
    MsgBox "Ocorreu o seguinte erro na rotina mAlterarArea: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub

Private Sub mDadosArea()

    On Error GoTo ErrmDadosArea
    
    With lvwCadastroAuxArea.SelectedItem
        txtDescricaoArea.Text = .SubItems(1)
    End With
    
    Exit Sub
    
ErrmDadosArea:

    MsgBox "Ocorreu o seguinte erro na rotina mDadosArea: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub



Private Sub mExcluirArea()

    Dim strSql            As String
    Dim lngLinhasAfetadas As Long
    
    On Error GoTo ErrmExcluirArea
    
    Screen.MousePointer = vbHourglass
    
    strSql = "SP_EXCLUIR_Area " & CInt(lvwCadastroAuxArea.SelectedItem.Text)
             
    dbConexaoSMR.Execute strSql, lngLinhasAfetadas
    
    If lngLinhasAfetadas > 0 Then
        MsgBox "Area excluída com sucesso ! ", vbInformation + vbOKOnly, App.Title
    End If
    
    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadAuxArea.Visible = False
    fraListaCadAuxArea.Visible = True
    
    'Preencher lista
    Call mPreencheListaArea
    
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrmExcluirArea:

    Screen.MousePointer = vbNormal
    
    MsgBox "Ocorreu o seguinte erro na rotina mExcluirArea: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub

Private Sub mHabilitarDesabilitarCamposCadastroArea(ByVal pAcao As String)

    fraDadosCadAuxArea.Tag = pAcao

    If UCase(Trim(pAcao)) = "I" Then
        
        txtDescricaoArea.Enabled = True
        
        txtDescricaoArea.SetFocus
    
    ElseIf UCase(Trim(pAcao)) = "A" Then
        
        txtDescricaoArea.Enabled = True
        
    End If
    
End Sub

Private Sub mIncluirArea()

    Dim strSql                  As String
    Dim lngCodigoArea     As Long
    Dim lngLinhasAfetadas       As Long
    Dim itmAreaInserida   As ListItem
    Dim cmdArea           As ADODB.Command
    
    On Error GoTo ErrmIncluirArea
    
    Screen.MousePointer = vbHourglass
    
    Set cmdArea = New ADODB.Command
    
    With cmdArea
    
        .ActiveConnection = dbConexaoSMR
        .CommandType = adCmdStoredProc
        .CommandText = "SP_INCLUIR_Area"
        
        .Parameters.Refresh
        .Parameters(1).Value = Trim(txtDescricaoArea.Text)
        
        .Execute
        
        lngCodigoArea = .Parameters(2).Value
        
    End With
        
    If lngCodigoArea > 0 Then
        MsgBox "Area incluída com sucesso ! ", vbInformation + vbOKOnly, App.Title
    End If
    
    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadAuxArea.Visible = False
    fraListaCadAuxArea.Visible = True
    
    'Adicionando na lista
    Set itmAreaInserida = lvwCadastroAuxArea.ListItems.Add(, , Format(lngCodigoArea, "000"))
    itmAreaInserida.SubItems(1) = txtDescricaoArea.Text
    
    itmAreaInserida.Selected = True
    itmAreaInserida.EnsureVisible
        
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrmIncluirArea:

    Screen.MousePointer = vbNormal
    
    MsgBox "Ocorreu o seguinte erro na rotina mIncluirArea: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub

Private Sub mLimparDadosArea()

    txtDescricaoArea.Text = ""
    
End Sub


Private Sub mMontaCabecalhoListaArea()

    'Criando os cabeçalhos das colunas da lista
    With lvwCadastroAuxArea
        .ColumnHeaders.Add , , "Código", 1000
        .ColumnHeaders.Add , , "Descrição", 9370
    End With
    
End Sub

Private Sub mPreencheListaArea()

    Dim rsArea   As ADODB.Recordset
    Dim itmArea   As ListItem
    Dim strSql          As String
    
    On Error GoTo ErrmPreencheListaArea
    
    Screen.MousePointer = vbHourglass
    
    With lvwCadastroAuxArea
    
        .ListItems.Clear
        .ColumnHeaders.Clear
        
        'Monta o cabeçalho da lista
        Call mMontaCabecalhoListaArea
        
        strSql = "EXECUTE SP_LISTAR_AREAS"
        
        Set rsArea = New ADODB.Recordset
        rsArea.Open strSql, dbConexaoSMR, adOpenStatic, adLockReadOnly, adCmdText
        
        Do While Not rsArea.EOF
        
            Set itmArea = .ListItems.Add(, , Format(rsArea.Fields("Cod_Area").Value, "000"))
            itmArea.SubItems(1) = rsArea.Fields("Desc_Area").Value
              
            rsArea.MoveNext
            
        Loop
        
    End With
    
    rsArea.Close
    Set rsArea = Nothing
    
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrmPreencheListaArea:

    Screen.MousePointer = vbNormal
    
    MsgBox "Ocorreu o seguinte erro na rotina mPreencheListaArea: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title
           
    If rsArea.State = adStateOpen Then
        rsArea.Close
        Set rsArea = Nothing
    End If

End Sub

Private Sub mTratarTelaAlterarArea()

    On Error GoTo ErrmTratarTelaAlterarArea
    
    'Desabilita os botões da ToolBar
    Call mDesabilitarBotoesToolBarArea
    
    'Coloca o frame que contém os dados sobreposto ao frame de lista
    fraDadosCadAuxArea.Visible = True
    fraListaCadAuxArea.Visible = False
        
    Exit Sub
    
ErrmTratarTelaAlterarArea:

    MsgBox "Ocorreu o seguinte erro na rotina mTratarTelaAlterarArea: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub

Private Sub mTratarTelaExcluirArea()

    On Error GoTo ErrmTratarTelaExcluirArea
    
    'Coloca o frame que contém os dados sobreposto ao frame de lista
    fraDadosCadAuxArea.Visible = True
    fraListaCadAuxArea.Visible = False
    
    Exit Sub
    
ErrmTratarTelaExcluirArea:

    MsgBox "Ocorreu o seguinte erro na rotina mTratarTelaExcluirArea: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub


Private Sub mTratarTelaIncluirArea()

    On Error GoTo ErrmTratarTelaIncluirArea
    
    'Desabilita os botões da ToolBar
    Call mDesabilitarBotoesToolBarArea
    
    'Limpa dados
    Call mLimparDadosArea
    
    'Coloca o frame que contém os dados sobreposto ao frame de lista
    fraDadosCadAuxArea.Visible = True
    fraListaCadAuxArea.Visible = False
    
    Exit Sub
    
ErrmTratarTelaIncluirArea:

    MsgBox "Ocorreu o seguinte erro na rotina mTratarTelaIncluirArea: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title
    
End Sub

Private Function mValidarCamposArea() As Boolean

    On Error GoTo ErrmValidarCamposArea
    
    If Trim(txtDescricaoArea.Text) = "" Then
        MsgBox "Descrição da Area não foi preenchida !", vbInformation + vbOKOnly, App.Title
        txtDescricaoArea.SetFocus
        mValidarCamposArea = False
        Exit Function
    End If
        
    mValidarCamposArea = True
        
    Exit Function
    
ErrmValidarCamposArea:

    MsgBox "Ocorreu o seguinte erro na rotina mValidarCamposArea: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title
    
End Function

Private Sub cmdCancelar_Click()

    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadAuxArea.Visible = False
    fraListaCadAuxArea.Visible = True

    'Habilita os botões da ToolBar
    Call mHabilitarBotoesToolBarArea

End Sub

Private Sub cmdGravar_Click()
    
    On Error GoTo ErrcmdGravar_Click
    
    If mValidarCamposArea() Then
    
        If fraDadosCadAuxArea.Tag = "I" Then
        
            'Inclusao
            Call mIncluirArea
        
        ElseIf fraDadosCadAuxArea.Tag = "A" Then
            
            'Alteração
            Call mAlterarArea
        
        End If
        
        'Habilita os botões da ToolBar
        Call mHabilitarBotoesToolBarArea
        
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
    Call mHabilitarBotoesToolBarArea
        
    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadAuxArea.Visible = False
    fraListaCadAuxArea.Visible = True

    'Preenche a lista
    Call mPreencheListaArea
    
End Sub









Private Sub Form_Unload(Cancel As Integer)

    frmPrincipal.StatusBar.Panels(1).Text = ""
    
End Sub


Private Sub lvwCadastroAuxArea_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    With lvwCadastroAuxArea
    
        .SortKey = ColumnHeader.Index - 1
        
        If .SortOrder = lvwAscending Then
           .SortOrder = lvwDescending
        Else
           .SortOrder = lvwAscending
        End If
        
        .Sorted = True
        
    End With

End Sub


Private Sub tlbCadastroAuxArea_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
    
        Case 1 'Incluir
        
            Call mTratarTelaIncluirArea
            Call mHabilitarDesabilitarCamposCadastroArea("I")
                    
        Case 2 'Alterar
        
            Call mTratarTelaAlterarArea
            Call mDadosArea
            Call mHabilitarDesabilitarCamposCadastroArea("A")
                        
        Case 3 'Excluir
        
            If MsgBox("Deseja realmente excluir este registro ?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                Call mExcluirArea
            End If
            
        Case 5 'Sair
            
            Unload Me
            
    End Select

End Sub


