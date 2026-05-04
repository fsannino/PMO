VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCadAtualizarPercSanSin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Atualizações de Tarefas"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   12675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.Toolbar tlbCadastroAtulTar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12675
      _ExtentX        =   22357
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
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   10830
         Top             =   105
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
               Picture         =   "frmCadAtualizarPercSanSin.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAtualizarPercSanSin.frx":0112
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAtualizarPercSanSin.frx":0224
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAtualizarPercSanSin.frx":038E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAtualizarPercSanSin.frx":04A0
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadAtualizarPercSanSin.frx":08F2
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraListaCadAtulTar 
      Caption         =   "Lista de Atualizações de Tarefas Cadastradas"
      Height          =   4530
      Left            =   45
      TabIndex        =   7
      Top             =   525
      Width           =   12570
      Begin MSComctlLib.ListView lvwCadastroAtulTar 
         Height          =   4185
         Left            =   45
         TabIndex        =   8
         Top             =   255
         Width           =   12420
         _ExtentX        =   21908
         _ExtentY        =   7382
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
   Begin VB.Frame fraDadosCadAtulTar 
      Caption         =   "Dados da Atualização da Tarefa"
      Height          =   4530
      Left            =   45
      TabIndex        =   9
      Top             =   525
      Width           =   12570
      Begin VB.Frame frmDestino 
         Caption         =   "Destino"
         Height          =   1680
         Index           =   1
         Left            =   390
         TabIndex        =   11
         Top             =   2145
         Width           =   11745
         Begin VB.ComboBox cmbTarefaDes 
            Height          =   315
            Left            =   2430
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1035
            Width           =   7860
         End
         Begin VB.ComboBox cmbProjetoDes 
            Height          =   315
            Left            =   2415
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   390
            Width           =   7875
         End
         Begin VB.Label lblTarefaDes 
            AutoSize        =   -1  'True
            Caption         =   "Escolha a Tarefa:"
            Height          =   195
            Index           =   0
            Left            =   990
            TabIndex        =   15
            Top             =   1095
            Width           =   1260
         End
         Begin VB.Label lblProjetoDes 
            AutoSize        =   -1  'True
            Caption         =   "Escolha o projeto:"
            Height          =   195
            Index           =   1
            Left            =   990
            TabIndex        =   14
            Top             =   465
            Width           =   1275
         End
      End
      Begin VB.Frame frmOrigem 
         Caption         =   "Origem"
         Height          =   1680
         Index           =   0
         Left            =   375
         TabIndex        =   10
         Top             =   300
         Width           =   11760
         Begin VB.ComboBox cmbTarefaOrg 
            Height          =   315
            Left            =   2430
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   975
            Width           =   7845
         End
         Begin VB.ComboBox cmbProjetoOrg 
            Height          =   315
            Left            =   2415
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   375
            Width           =   7875
         End
         Begin VB.Label lblTarefaOrg 
            AutoSize        =   -1  'True
            Caption         =   "Escolha a Tarefa:"
            Height          =   195
            Index           =   1
            Left            =   990
            TabIndex        =   13
            Top             =   1080
            Width           =   1260
         End
         Begin VB.Label lblProjetoOrg 
            AutoSize        =   -1  'True
            Caption         =   "Escolha o projeto:"
            Height          =   195
            Index           =   0
            Left            =   990
            TabIndex        =   12
            Top             =   450
            Width           =   1275
         End
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   11100
         TabIndex        =   6
         Top             =   3990
         Width           =   1020
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "Gravar"
         Height          =   375
         Left            =   9885
         TabIndex        =   5
         Top             =   3990
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frmCadAtualizarPercSanSin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mHabilitarBotoesToolBarAtulTar()

    With tlbCadastroAtulTar
    
        .Buttons(1).Enabled = True
        .Buttons(2).Enabled = False
        .Buttons(3).Enabled = True
        
    End With
    
End Sub

Private Sub mDesabilitarBotoesToolBarAtulTar()

    With tlbCadastroAtulTar
    
        .Buttons(1).Enabled = False
        .Buttons(2).Enabled = False
        .Buttons(3).Enabled = False
        
    End With
    
End Sub

Private Function mValidarCamposAtulTar() As Boolean

    On Error GoTo ErrmValidarCamposAtulTar
    
    If Trim(cmbProjetoOrg.Text) = "" Then
        MsgBox "Projeto origem não foi preenchido !", vbInformation + vbOKOnly, App.Title
        cmbProjetoOrg.SetFocus
        mValidarCamposAtulTar = False
        Exit Function
    End If

    If Trim(cmbTarefaOrg.Text) = "" Then
        MsgBox "Tarefa origem não foi preenchida !", vbInformation + vbOKOnly, App.Title
        cmbTarefaOrg.SetFocus
        mValidarCamposAtulTar = False
        Exit Function
    End If

    If Trim(cmbProjetoDes.Text) = "" Then
        MsgBox "Projeto destino não foi preenchido !", vbInformation + vbOKOnly, App.Title
        cmbProjetoDes.SetFocus
        mValidarCamposAtulTar = False
        Exit Function
    End If

    If Trim(cmbTarefaDes.Text) = "" Then
        MsgBox "Tarefa destino não foi preenchida !", vbInformation + vbOKOnly, App.Title
        cmbTarefaDes.SetFocus
        mValidarCamposAtulTar = False
        Exit Function
    End If
    
    mValidarCamposAtulTar = True
        
    Exit Function
    
ErrmValidarCamposAtulTar:

    MsgBox "Ocorreu o seguinte erro na rotina mValidarCamposAtulTar: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title
    
End Function

Private Sub mIncluirAtulTar()

    Dim strSQL                      As String
    Dim lngLinhasAfetadas           As Long
    Dim itmAtulTarInserido          As ListItem
    Dim rsAtualizaTarExiste         As ADODB.Recordset
    
    Dim strCodProjOrg               As String
    Dim strCodTarOrg                As String
    Dim strCodProjDes               As String
    Dim strCodTarDes                As String
    
    On Error GoTo ErrmIncluirAtulTar
    
    Screen.MousePointer = vbHourglass
    
    strCodProjOrg = Mid(cmbProjetoOrg.Text, 1, (InStr(1, cmbProjetoOrg.Text, "-") - 2))
    strCodTarOrg = Mid(cmbTarefaOrg.Text, 1, (InStr(1, cmbTarefaOrg.Text, "-") - 2))
    strCodProjDes = Mid(cmbProjetoDes.Text, 1, (InStr(1, cmbProjetoDes.Text, "-") - 2))
    strCodTarDes = Mid(cmbTarefaDes.Text, 1, (InStr(1, cmbTarefaDes.Text, "-") - 2))
    
    strSQL = "EXECUTE SP_LISTAR_ATUALIZA_PERC_SAN_SIN " & strCodProjOrg & ", " & strCodTarOrg & ", " & strCodProjDes & ", " & strCodTarDes

    Set rsAtualizaTarExiste = New ADODB.Recordset
    rsAtualizaTarExiste.CursorLocation = adUseClient
    rsAtualizaTarExiste.Open strSQL, dbConexaoSMR, adOpenStatic, adLockReadOnly, adCmdText


    If rsAtualizaTarExiste.RecordCount = 0 Then

        strSQL = "SP_INCLUIR_ATUALIZA_PERC_SAN_SIN " & strCodProjOrg & ", " & strCodTarOrg & ", " & strCodProjDes & ", " & strCodTarDes

        dbConexaoSMR.Execute strSQL, lngLinhasAfetadas

        If lngLinhasAfetadas > 0 Then
            MsgBox "Atualização de tarefa incluída com sucesso ! ", vbInformation + vbOKOnly, App.Title
        End If

        'Coloca o frame que contém a lista sobreposto ao frame de dados
        fraDadosCadAtulTar.Visible = False
        fraListaCadAtulTar.Visible = True

        Call mPreencheListaAtulTar

    Else

        MsgBox "Atualização de tarefa já existe ! ", vbInformation + vbOKOnly, App.Title
        cmbProjetoOrg.SetFocus

    End If
    
        
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrmIncluirAtulTar:

    Screen.MousePointer = vbNormal
    
    MsgBox "Ocorreu o seguinte erro na rotina mIncluirAtulTar: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub

Private Sub mExcluirAtulTar()

    Dim strSQL                      As String
    Dim lngLinhasAfetadas           As Long
    
    Dim strCodProjOrg               As String
    Dim strCodTarOrg                As String
    Dim strCodProjDes               As String
    Dim strCodTarDes                As String
    
    On Error GoTo ErrmExcluirAtulTar
    
    Screen.MousePointer = vbHourglass
        
    strCodProjOrg = lvwCadastroAtulTar.SelectedItem.Text
    strCodTarOrg = lvwCadastroAtulTar.SelectedItem.SubItems(5)
    strCodProjDes = lvwCadastroAtulTar.SelectedItem.SubItems(6)
    strCodTarDes = lvwCadastroAtulTar.SelectedItem.SubItems(7)
        
    strSQL = "SP_EXCLUIR_ATUALIZA_PERC_SAN_SIN " & strCodProjOrg & ", " & strCodTarOrg & ", " & strCodProjDes & ", " & strCodTarDes
             
    dbConexaoSMR.Execute strSQL, lngLinhasAfetadas
    
    If lngLinhasAfetadas > 0 Then
        MsgBox "Atualização de tarefa excluída com sucesso ! ", vbInformation + vbOKOnly, App.Title
    End If
    
    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadAtulTar.Visible = False
    fraListaCadAtulTar.Visible = True
    
    'Preencher lista de usuários
    Call mPreencheListaAtulTar
    
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrmExcluirAtulTar:

    Screen.MousePointer = vbNormal
    
    MsgBox "Ocorreu o seguinte erro na rotina mExcluirAtulTar: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub



Private Sub mTratarTelaIncluirAtulTar()

    On Error GoTo ErrmTratarTelaIncluirAtulTar
    
    'Desabilita os botões da ToolBar
    Call mDesabilitarBotoesToolBarAtulTar
    
    'Coloca o frame que contém os dados sobreposto ao frame de lista
    fraDadosCadAtulTar.Visible = True
    fraListaCadAtulTar.Visible = False
    
    Call mLimparDadosAtulTar
    
    Exit Sub
    
ErrmTratarTelaIncluirAtulTar:

    MsgBox "Ocorreu o seguinte erro na rotina mTratarTelaIncluirAtulTar: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title
    
End Sub

Private Sub mTratarTelaExcluirAtulTar()

    On Error GoTo ErrmTratarTelaExcluirAtulTar
    
    'Coloca o frame que contém os dados sobreposto ao frame de lista
    fraDadosCadAtulTar.Visible = True
    fraListaCadAtulTar.Visible = False
    
    Exit Sub
    
ErrmTratarTelaExcluirAtulTar:

    MsgBox "Ocorreu o seguinte erro na rotina mTratarTelaExcluirAtulTar: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub

Private Sub mPreencheListaAtulTar()

    Dim rsAtulTar   As ADODB.Recordset
    Dim itmAtulTar   As ListItem
    Dim strSQL          As String
    
    On Error GoTo ErrmPreencheListaAtulTar
    
    Screen.MousePointer = vbHourglass
    
    With lvwCadastroAtulTar
    
        .ListItems.Clear
        .ColumnHeaders.Clear
    
        'Monta o cabeçalho da lista
        Call mMontaCabecalhoListaAtulTar
        
        strSQL = "EXECUTE SP_LISTAR_ATUALIZA_PERC_SAN_SIN"
        
        Set rsAtulTar = New ADODB.Recordset
        rsAtulTar.Open strSQL, dbConexaoSMR, adOpenStatic, adLockReadOnly, adCmdText
        
        Do While Not rsAtulTar.EOF
        
            Set itmAtulTar = .ListItems.Add(, , rsAtulTar.Fields("Cod_Proj_Orig").Value)
            itmAtulTar.SubItems(1) = rsAtulTar.Fields("Proj_Org").Value
            itmAtulTar.SubItems(2) = rsAtulTar.Fields("Tar_Org").Value
            itmAtulTar.SubItems(3) = rsAtulTar.Fields("Proj_Des").Value
            itmAtulTar.SubItems(4) = rsAtulTar.Fields("Tar_Des").Value
            itmAtulTar.SubItems(5) = rsAtulTar.Fields("Cod_UID_Orig").Value
            itmAtulTar.SubItems(6) = rsAtulTar.Fields("Cod_Proj_Dest").Value
            itmAtulTar.SubItems(7) = rsAtulTar.Fields("Cod_UID_Dest").Value
              
            rsAtulTar.MoveNext
            
        Loop
    
    End With
    
    rsAtulTar.Close
    Set rsAtulTar = Nothing
    
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrmPreencheListaAtulTar:

    Screen.MousePointer = vbNormal
    
    MsgBox "Ocorreu o seguinte erro na rotina mPreencheListaAtulTar: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title
           
    If rsAtulTar.State = adStateOpen Then
        rsAtulTar.Close
        Set rsAtulTar = Nothing
    End If

End Sub

Private Sub mMontaCabecalhoListaAtulTar()

    'Criando os cabeçalhos das colunas da lista de usuários
    With lvwCadastroAtulTar
        .ColumnHeaders.Add , , "", 0
        .ColumnHeaders.Add , , "Projeto Origem", 2500
        .ColumnHeaders.Add , , "Tarefa Origem", 6200
        .ColumnHeaders.Add , , "Projeto Destino", 2500
        .ColumnHeaders.Add , , "Tarefa Destino", 6200
        .ColumnHeaders.Add , , "UID Org.", 0
        .ColumnHeaders.Add , , "Cod.Proj.Des.", 0
        .ColumnHeaders.Add , , "UID Des.", 0
    End With
    
End Sub

Private Sub cmbProjetoDes_Click()
Dim strCodPrjetoDes As String
    
    cmbTarefaDes.Clear
    
    If Trim(cmbProjetoDes.Text) <> "" Then
        strCodPrjetoDes = Mid(cmbProjetoDes.Text, 1, (InStr(1, cmbProjetoDes.Text, "-") - 2))
        Call mPreencherComboTarefasDes(strCodPrjetoDes)
    End If
    
End Sub

Private Sub cmbProjetoOrg_Click()
Dim strCodPrjetoOrg As String
    
    cmbTarefaOrg.Clear
    
    If Trim(cmbProjetoOrg.Text) <> "" Then
        strCodPrjetoOrg = Mid(cmbProjetoOrg.Text, 1, (InStr(1, cmbProjetoOrg.Text, "-") - 2))
        Call mPreencherComboTarefasOrg(strCodPrjetoOrg)
    End If
End Sub

Private Sub cmdCancelar_Click()

    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadAtulTar.Visible = False
    fraListaCadAtulTar.Visible = True

    'Habilita os botões da ToolBar
    Call mHabilitarBotoesToolBarAtulTar

End Sub

Private Sub cmdGravar_Click()

    On Error GoTo ErrcmdGravar_Click
    
    If mValidarCamposAtulTar() Then
            
        'Inclusao
        Call mIncluirAtulTar
        
        'Habilita os botões da ToolBar
        Call mHabilitarBotoesToolBarAtulTar
        
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
    Call mHabilitarBotoesToolBarAtulTar

    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadAtulTar.Visible = False
    fraListaCadAtulTar.Visible = True

    'Preenche a lista de Projetos PMO
    Call mPreencheListaAtulTar
    Call mPreencherComboProjetos
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    frmPrincipal.StatusBar.Panels(1).Text = ""
    
End Sub


Private Sub lvwCadastroAtulTar_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    With lvwCadastroAtulTar
    
        .SortKey = ColumnHeader.Index - 1
        
        If .SortOrder = lvwAscending Then
           .SortOrder = lvwDescending
        Else
           .SortOrder = lvwAscending
        End If
        
        .Sorted = True
        
    End With

End Sub


Private Sub tlbCadastroAtulTar_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
    
        Case 1 'Incluir
        
            Call mTratarTelaIncluirAtulTar
                        
        Case 3 'Excluir
        
            If MsgBox("Deseja realmente excluir este registro ?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                Call mExcluirAtulTar
            End If
            
        Case 5 'Sair
            
            Unload Me
            
    End Select

End Sub


Private Sub mPreencherComboProjetos()

    Dim rsProjetos  As ADODB.Recordset
    Dim strSQL      As String

    On Error GoTo ErrmPreencherComboProjetos
    
    strSQL = "EXECUTE SP_LISTAR_PROJETOS_PROJECT"
    
    Set rsProjetos = New ADODB.Recordset
    rsProjetos.Open strSQL, dbConexaoSMR, adOpenStatic, adLockReadOnly, adCmdText
    
    cmbProjetoOrg.AddItem ""
    cmbProjetoDes.AddItem ""
    
    Do While Not rsProjetos.EOF
    
        cmbProjetoOrg.AddItem rsProjetos.Fields("PROJ_ID").Value & " - " & rsProjetos.Fields("PROJ_NAME").Value
        cmbProjetoDes.AddItem rsProjetos.Fields("PROJ_ID").Value & " - " & rsProjetos.Fields("PROJ_NAME").Value
        
        rsProjetos.MoveNext
        
    Loop
    
    cmbProjetoOrg.ListIndex = 0
    cmbProjetoDes.ListIndex = 0
    
    rsProjetos.Close
    Set rsProjetos = Nothing
    
    Exit Sub
    
ErrmPreencherComboProjetos:

    MsgBox "Ocorreu o seguinte erro na rotina mPreencherComboProjetos: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title
           
    If rsProjetos.State = adStateOpen Then
        rsProjetos.Close
        Set rsProjetos = Nothing
    End If
    
End Sub


Private Sub mPreencherComboTarefasOrg(strCodProjeto As String)

    Dim rsTarefas  As ADODB.Recordset
    Dim strSQL      As String

    On Error GoTo ErrmPreencherComboTarefasOrg
    
    strSQL = "EXECUTE SP_LISTAR_TAREFAS " & strCodProjeto & ""
    
    Set rsTarefas = New ADODB.Recordset
    rsTarefas.Open strSQL, dbConexaoSMR, adOpenStatic, adLockReadOnly, adCmdText
    
    cmbTarefaOrg.Clear
    
    cmbTarefaOrg.AddItem ""
    
    Do While Not rsTarefas.EOF
    
        cmbTarefaOrg.AddItem rsTarefas.Fields("TASK_UID").Value & " - " & rsTarefas.Fields("TASK_NAME").Value
        
        rsTarefas.MoveNext
        
    Loop
    
    cmbTarefaOrg.ListIndex = 0
    
    rsTarefas.Close
    Set rsTarefas = Nothing
    
    Exit Sub
    
ErrmPreencherComboTarefasOrg:
    
    MsgBox "Ocorreu o seguinte erro na rotina mPreencherComboTarefasOrg: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title
           
    If rsTarefas.State = adStateOpen Then
        rsTarefas.Close
        Set rsTarefas = Nothing
    End If
    
End Sub


Private Sub mPreencherComboTarefasDes(strCodProjeto As String)

    Dim rsTarefas  As ADODB.Recordset
    Dim strSQL      As String

    On Error GoTo ErrmPreencherComboTarefasDes
    
    strSQL = "EXECUTE SP_LISTAR_TAREFAS " & strCodProjeto & ""
    
    Set rsTarefas = New ADODB.Recordset
    rsTarefas.Open strSQL, dbConexaoSMR, adOpenStatic, adLockReadOnly, adCmdText
    
    cmbTarefaDes.AddItem ""
    
    Do While Not rsTarefas.EOF
    
        cmbTarefaDes.AddItem rsTarefas.Fields("TASK_UID").Value & " - " & rsTarefas.Fields("TASK_NAME").Value
        
        rsTarefas.MoveNext
        
    Loop
    
    cmbTarefaDes.ListIndex = 0
    
    rsTarefas.Close
    Set rsTarefas = Nothing
    
    Exit Sub
    
ErrmPreencherComboTarefasDes:
    
    MsgBox "Ocorreu o seguinte erro na rotina mPreencherComboTarefasDes: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title
           
    If rsTarefas.State = adStateOpen Then
        rsTarefas.Close
        Set rsTarefas = Nothing
    End If
    
End Sub


Private Sub mLimparDadosAtulTar()

cmbProjetoOrg.ListIndex = 0
cmbProjetoDes.ListIndex = 0

cmbTarefaDes.Clear
cmbTarefaOrg.Clear

End Sub

