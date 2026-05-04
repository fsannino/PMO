VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCadConfigFechamento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuração de Fechamentos "
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10620
   Icon            =   "frmCadConfigFechamento.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.Toolbar tlbCadastroConfFech 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   22
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
   Begin VB.Frame fraListaCadConfFech 
      Caption         =   "Lista Configuração do Fechamento"
      Height          =   4770
      Left            =   75
      TabIndex        =   23
      Top             =   525
      Width           =   10500
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
               Picture         =   "frmCadConfigFechamento.frx":0442
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadConfigFechamento.frx":0554
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadConfigFechamento.frx":0666
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadConfigFechamento.frx":07D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadConfigFechamento.frx":08E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadConfigFechamento.frx":0D34
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvwCadastroConfFech 
         Height          =   4425
         Left            =   60
         TabIndex        =   24
         Top             =   255
         Width           =   10365
         _ExtentX        =   18283
         _ExtentY        =   7805
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
   Begin VB.Frame fraDadosCadConfFech 
      Caption         =   "Configuração do Fechamento"
      Height          =   4755
      Left            =   75
      TabIndex        =   25
      Top             =   540
      Width           =   10500
      Begin VB.Frame Frame9 
         Caption         =   "Sigla do Projeto"
         Height          =   1575
         Left            =   240
         TabIndex        =   50
         Top             =   285
         Width           =   2220
         Begin VB.TextBox txtSiglaProj 
            Height          =   450
            Left            =   285
            TabIndex        =   0
            Top             =   585
            Width           =   1560
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Domingo"
         Height          =   1575
         Left            =   8025
         TabIndex        =   47
         Top             =   2055
         Width           =   2220
         Begin VB.CheckBox chkFechado_Dom 
            Caption         =   "Fechado"
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   1890
         End
         Begin MSMask.MaskEdBox mskInicio_Dom 
            Height          =   360
            Left            =   930
            TabIndex        =   20
            Top             =   630
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   635
            _Version        =   393216
            MaxLength       =   8
            Format          =   "hh:mm:ss"
            Mask            =   "##:##:##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskFim_Dom 
            Height          =   360
            Left            =   930
            TabIndex        =   21
            Top             =   1080
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   635
            _Version        =   393216
            MaxLength       =   8
            Format          =   "hh:mm:ss"
            Mask            =   "##:##:##"
            PromptChar      =   " "
         End
         Begin VB.Label Label14 
            Caption         =   "HR. Inicio:"
            Height          =   255
            Left            =   90
            TabIndex        =   49
            Top             =   705
            Width           =   855
         End
         Begin VB.Label Label13 
            Caption         =   "HR.Fim:"
            Height          =   255
            Left            =   90
            TabIndex        =   48
            Top             =   1140
            Width           =   855
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Sabado"
         Height          =   1575
         Left            =   5385
         TabIndex        =   44
         Top             =   2055
         Width           =   2220
         Begin VB.CheckBox chkFechado_Sab 
            Caption         =   "Fechado"
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   1890
         End
         Begin MSMask.MaskEdBox mskInicio_Sab 
            Height          =   360
            Left            =   930
            TabIndex        =   17
            Top             =   630
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   635
            _Version        =   393216
            MaxLength       =   8
            Format          =   "hh:mm:ss"
            Mask            =   "##:##:##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskFim_Sab 
            Height          =   360
            Left            =   930
            TabIndex        =   18
            Top             =   1080
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   635
            _Version        =   393216
            MaxLength       =   8
            Format          =   "hh:mm:ss"
            Mask            =   "##:##:##"
            PromptChar      =   " "
         End
         Begin VB.Label Label12 
            Caption         =   "HR.Fim:"
            Height          =   255
            Left            =   90
            TabIndex        =   46
            Top             =   1140
            Width           =   855
         End
         Begin VB.Label Label11 
            Caption         =   "HR. Inicio:"
            Height          =   255
            Left            =   90
            TabIndex        =   45
            Top             =   705
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Sexta"
         Height          =   1575
         Left            =   2790
         TabIndex        =   41
         Top             =   2055
         Width           =   2220
         Begin VB.CheckBox chkFechado_Sex 
            Caption         =   "Fechado"
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1890
         End
         Begin MSMask.MaskEdBox mskInicio_Sex 
            Height          =   360
            Left            =   930
            TabIndex        =   14
            Top             =   630
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   635
            _Version        =   393216
            MaxLength       =   8
            Format          =   "hh:mm:ss"
            Mask            =   "##:##:##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskFim_Sex 
            Height          =   360
            Left            =   930
            TabIndex        =   15
            Top             =   1080
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   635
            _Version        =   393216
            MaxLength       =   8
            Format          =   "hh:mm:ss"
            Mask            =   "##:##:##"
            PromptChar      =   " "
         End
         Begin VB.Label Label10 
            Caption         =   "HR. Inicio:"
            Height          =   255
            Left            =   90
            TabIndex        =   43
            Top             =   705
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "HR.Fim:"
            Height          =   255
            Left            =   90
            TabIndex        =   42
            Top             =   1140
            Width           =   855
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Quinta"
         Height          =   1575
         Left            =   255
         TabIndex        =   38
         Top             =   2055
         Width           =   2220
         Begin VB.CheckBox chkFechado_Qui 
            Caption         =   "Fechado"
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   1890
         End
         Begin MSMask.MaskEdBox mskInicio_Qui 
            Height          =   360
            Left            =   930
            TabIndex        =   11
            Top             =   630
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   635
            _Version        =   393216
            MaxLength       =   8
            Format          =   "hh:mm:ss"
            Mask            =   "##:##:##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskFim_Qui 
            Height          =   360
            Left            =   930
            TabIndex        =   12
            Top             =   1080
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   635
            _Version        =   393216
            MaxLength       =   8
            Format          =   "hh:mm:ss"
            Mask            =   "##:##:##"
            PromptChar      =   " "
         End
         Begin VB.Label Label8 
            Caption         =   "HR.Fim:"
            Height          =   255
            Left            =   90
            TabIndex        =   40
            Top             =   1140
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "HR. Inicio:"
            Height          =   255
            Left            =   90
            TabIndex        =   39
            Top             =   705
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Quarta"
         Height          =   1575
         Left            =   8010
         TabIndex        =   35
         Top             =   285
         Width           =   2220
         Begin VB.CheckBox chkFechado_Qua 
            Caption         =   "Fechado"
            Height          =   375
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   1890
         End
         Begin MSMask.MaskEdBox mskInicio_Qua 
            Height          =   360
            Left            =   930
            TabIndex        =   8
            Top             =   630
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   635
            _Version        =   393216
            MaxLength       =   8
            Format          =   "hh:mm:ss"
            Mask            =   "##:##:##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskFim_Qua 
            Height          =   360
            Left            =   930
            TabIndex        =   9
            Top             =   1080
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   635
            _Version        =   393216
            MaxLength       =   8
            Format          =   "hh:mm:ss"
            Mask            =   "##:##:##"
            PromptChar      =   " "
         End
         Begin VB.Label Label6 
            Caption         =   "HR.Fim:"
            Height          =   255
            Left            =   90
            TabIndex        =   37
            Top             =   1140
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "HR. Inicio:"
            Height          =   255
            Left            =   90
            TabIndex        =   36
            Top             =   705
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Terça"
         Height          =   1575
         Left            =   5370
         TabIndex        =   32
         Top             =   285
         Width           =   2220
         Begin VB.CheckBox chkFechado_Ter 
            Caption         =   "Fechado"
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   1890
         End
         Begin MSMask.MaskEdBox mskInicio_Ter 
            Height          =   360
            Left            =   930
            TabIndex        =   5
            Top             =   630
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   635
            _Version        =   393216
            MaxLength       =   8
            Format          =   "hh:mm:ss"
            Mask            =   "##:##:##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskFim_Ter 
            Height          =   360
            Left            =   930
            TabIndex        =   6
            Top             =   1080
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   635
            _Version        =   393216
            MaxLength       =   8
            Format          =   "hh:mm:ss"
            Mask            =   "##:##:##"
            PromptChar      =   " "
         End
         Begin VB.Label Label4 
            Caption         =   "HR. Inicio:"
            Height          =   255
            Left            =   90
            TabIndex        =   34
            Top             =   705
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "HR.Fim:"
            Height          =   255
            Left            =   90
            TabIndex        =   33
            Top             =   1140
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Segunda"
         Height          =   1575
         Left            =   2790
         TabIndex        =   29
         Top             =   285
         Width           =   2220
         Begin MSMask.MaskEdBox mskInicio_Seg 
            Height          =   360
            Left            =   930
            TabIndex        =   2
            Top             =   630
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   635
            _Version        =   393216
            MaxLength       =   8
            Format          =   "hh:mm:ss"
            Mask            =   "##:##:##"
            PromptChar      =   " "
         End
         Begin VB.CheckBox chkFechado_Seg 
            Caption         =   "Fechado"
            Height          =   375
            Left            =   120
            TabIndex        =   1
            Top             =   240
            Width           =   1890
         End
         Begin MSMask.MaskEdBox mskFim_Seg 
            Height          =   360
            Left            =   930
            TabIndex        =   3
            Top             =   1080
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   635
            _Version        =   393216
            MaxLength       =   8
            Format          =   "hh:mm:ss"
            Mask            =   "##:##:##"
            PromptChar      =   " "
         End
         Begin VB.Label Label2 
            Caption         =   "HR.Fim:"
            Height          =   255
            Left            =   90
            TabIndex        =   31
            Top             =   1140
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "HR. Inicio:"
            Height          =   255
            Left            =   90
            TabIndex        =   30
            Top             =   705
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Height          =   795
         Left            =   255
         TabIndex        =   26
         Top             =   3780
         Width           =   9990
         Begin VB.CommandButton cmdGravar 
            Caption         =   "Gravar"
            Height          =   375
            Left            =   7605
            TabIndex        =   28
            Top             =   255
            Width           =   1020
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   8835
            TabIndex        =   27
            Top             =   255
            Width           =   1020
         End
      End
   End
End
Attribute VB_Name = "frmCadConfigFechamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub mDesabilitarBotoesToolBarConfFech()

    With tlbCadastroConfFech
    
        .Buttons(1).Enabled = False
        .Buttons(2).Enabled = False
        .Buttons(3).Enabled = False
        
    End With
    
End Sub

Private Sub mHabilitarBotoesToolBarConfFech()

    With tlbCadastroConfFech
    
        .Buttons(1).Enabled = False
        .Buttons(2).Enabled = True
        .Buttons(3).Enabled = False
        
    End With
    
End Sub

Private Sub mAlterarConfFech()

    Dim strSQL            As String
    Dim lngLinhasAfetadas As Long
    
    On Error GoTo ErrmAlterarConfFech
    
    Screen.MousePointer = vbHourglass
    
    strSQL = "SP_ALTERAR_CONFIG_FECHAMENTO '" & lvwCadastroConfFech.SelectedItem.Text & "', " & _
             "" & chkFechado_Seg.Value & ", '" & mskInicio_Seg.Text & "', '" & mskFim_Seg.Text & "', " & _
             "" & chkFechado_Ter.Value & ", '" & mskInicio_Ter.Text & "', '" & mskFim_Ter.Text & "', " & _
             "" & chkFechado_Qua.Value & ", '" & mskInicio_Qua.Text & "', '" & mskFim_Qua.Text & "', " & _
             "" & chkFechado_Qui.Value & ", '" & mskInicio_Qui.Text & "', '" & mskFim_Qui.Text & "', " & _
             "" & chkFechado_Sex.Value & ", '" & mskInicio_Sex.Text & "', '" & mskFim_Sex.Text & "', " & _
             "" & chkFechado_Sab.Value & ", '" & mskInicio_Sab.Text & "', '" & mskFim_Sab.Text & "', " & _
             "" & chkFechado_Dom.Value & ", '" & mskInicio_Dom.Text & "', '" & mskFim_Dom.Text & "' "
    
    dbConexaoSMR.Execute strSQL, lngLinhasAfetadas
    
    If lngLinhasAfetadas > 0 Then
        MsgBox "Configuração do fechamento alterada com sucesso ! ", vbInformation + vbOKOnly, App.Title
    End If
    
    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadConfFech.Visible = False
    fraListaCadConfFech.Visible = True
    
    'Alterando os dados na lista
    Call mPreencheListaConfFech
        
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrmAlterarConfFech:

    Screen.MousePointer = vbNormal
    
    MsgBox "Ocorreu o seguinte erro na rotina mAlterarConfFech: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub

Private Sub mDadosConfFech()

    On Error GoTo ErrmDadosConfFech
    
    With lvwCadastroConfFech.SelectedItem
        
        txtSiglaProj.Text = lvwCadastroConfFech.SelectedItem.Text
        
        'Segunda
        If .SubItems(1) = "Sim" Then
            chkFechado_Seg.Value = 1
        Else
            chkFechado_Seg.Value = 0
        End If
        
        mskInicio_Seg = .SubItems(2)
        mskFim_Seg = .SubItems(3)
        
        'Terça
        If .SubItems(4) = "Sim" Then
            chkFechado_Ter.Value = 1
        Else
            chkFechado_Ter.Value = 0
        End If
        
        mskInicio_Ter = .SubItems(5)
        mskFim_Ter = .SubItems(6)
        
        'Quarta
        If .SubItems(7) = "Sim" Then
            chkFechado_Qua.Value = 1
        Else
            chkFechado_Qua.Value = 0
        End If
        
        mskInicio_Qua = .SubItems(8)
        mskFim_Qua = .SubItems(9)
    
        'Quinta
        If .SubItems(10) = "Sim" Then
            chkFechado_Qui.Value = 1
        Else
            chkFechado_Qui.Value = 0
        End If
        
        mskInicio_Qui = .SubItems(11)
        mskFim_Qui = .SubItems(12)
    
        'Sexta
        If .SubItems(13) = "Sim" Then
            chkFechado_Sex.Value = 1
        Else
            chkFechado_Sex.Value = 0
        End If
        
        mskInicio_Sex = .SubItems(14)
        mskFim_Sex = .SubItems(15)
    
        'Sabado
        If .SubItems(16) = "Sim" Then
            chkFechado_Sab.Value = 1
        Else
            chkFechado_Sab.Value = 0
        End If
        
        mskInicio_Sab = .SubItems(17)
        mskFim_Sab = .SubItems(18)
    
        'Domingo
        If .SubItems(19) = "Sim" Then
            chkFechado_Dom.Value = 1
        Else
            chkFechado_Dom.Value = 0
        End If
        
        mskInicio_Dom = .SubItems(20)
        mskFim_Dom = .SubItems(21)
    
    
    End With
    
    Exit Sub
    
ErrmDadosConfFech:

    MsgBox "Ocorreu o seguinte erro na rotina mDadosConfFech: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub


Private Sub mHabilitarDesabilitarCamposCadastroConfFech(ByVal pAcao As String)

    fraDadosCadConfFech.Tag = pAcao
  
    If UCase(Trim(pAcao)) = "A" Then
        
        txtSiglaProj.Enabled = False
            
        'Segunda
        If chkFechado_Seg.Value = 0 Then
            mskInicio_Seg.Text = "00:00:00"
            mskInicio_Seg.Enabled = False
            mskFim_Seg.Text = "00:00:00"
            mskFim_Seg.Enabled = False
        Else
            mskInicio_Seg.Enabled = True
            mskFim_Seg.Enabled = True
        End If
    
        'Terça
        If chkFechado_Ter.Value = 0 Then
            mskInicio_Ter.Text = "00:00:00"
            mskInicio_Ter.Enabled = False
            mskFim_Ter.Text = "00:00:00"
            mskFim_Ter.Enabled = False
        Else
            mskInicio_Ter.Enabled = True
            mskFim_Ter.Enabled = True
        End If
    
        'Quarta
        If chkFechado_Qua.Value = 0 Then
            mskInicio_Qua.Text = "00:00:00"
            mskInicio_Qua.Enabled = False
            mskFim_Qua.Text = "00:00:00"
            mskFim_Qua.Enabled = False
        Else
            mskInicio_Qua.Enabled = True
            mskFim_Qua.Enabled = True
        End If
    
        'Quinta
        If chkFechado_Qui.Value = 0 Then
            mskInicio_Qui.Text = "00:00:00"
            mskInicio_Qui.Enabled = False
            mskFim_Qui.Text = "00:00:00"
            mskFim_Qui.Enabled = False
        Else
            mskInicio_Qui.Enabled = True
            mskFim_Qui.Enabled = True
        End If
    
        'Sexta
        If chkFechado_Sex.Value = 0 Then
            mskInicio_Sex.Text = "00:00:00"
            mskInicio_Sex.Enabled = False
            mskFim_Sex.Text = "00:00:00"
            mskFim_Sex.Enabled = False
        Else
            mskInicio_Sex.Enabled = True
            mskFim_Sex.Enabled = True
        End If
    
        'Sabado
        If chkFechado_Sab.Value = 0 Then
            mskInicio_Sab.Text = "00:00:00"
            mskInicio_Sab.Enabled = False
            mskFim_Sab.Text = "00:00:00"
            mskFim_Sab.Enabled = False
        Else
            mskInicio_Sab.Enabled = True
            mskFim_Sab.Enabled = True
        End If
    
        'Domingo
        If chkFechado_Dom.Value = 0 Then
            mskInicio_Dom.Text = "00:00:00"
            mskInicio_Dom.Enabled = False
            mskFim_Dom.Text = "00:00:00"
            mskFim_Dom.Enabled = False
        Else
            mskInicio_Dom.Enabled = True
            mskFim_Dom.Enabled = True
        End If
    
    End If
    
End Sub

Private Sub mMontaCabecalhoListaConfFech()

    'Criando os cabeçalhos das colunas da lista
    With lvwCadastroConfFech
        .ColumnHeaders.Add , , "Sigla Projeto", 1000
        .ColumnHeaders.Add , , "Segunda", 1000
        .ColumnHeaders.Add , , "Segunda HR.Inicio", 1000
        .ColumnHeaders.Add , , "Segunda HR.Fim", 1000
        .ColumnHeaders.Add , , "Terça", 1000
        .ColumnHeaders.Add , , "Terça HR.Inicio", 1000
        .ColumnHeaders.Add , , "Terça HR.Fim", 1000
        .ColumnHeaders.Add , , "Quarta", 1000
        .ColumnHeaders.Add , , "Quarta HR.Inicio", 1000
        .ColumnHeaders.Add , , "Quarta HR.Fim", 1000
        .ColumnHeaders.Add , , "Quinta", 1000
        .ColumnHeaders.Add , , "Quinta HR.Inicio", 1000
        .ColumnHeaders.Add , , "Quinta HR.Fim", 1000
        .ColumnHeaders.Add , , "Sexta", 1000
        .ColumnHeaders.Add , , "Sexta HR.Inicio", 1000
        .ColumnHeaders.Add , , "Sexta HR.Fim", 1000
        .ColumnHeaders.Add , , "Sabado", 1000
        .ColumnHeaders.Add , , "Sabado HR.Inicio", 1000
        .ColumnHeaders.Add , , "Sabado HR.Fim", 1000
        .ColumnHeaders.Add , , "Domingo", 1000
        .ColumnHeaders.Add , , "Domingo HR.Inicio", 1000
        .ColumnHeaders.Add , , "Domingo HR.Fim", 1000
    End With
    
End Sub

Private Sub mPreencheListaConfFech()

    Dim rsConfFech      As ADODB.Recordset
    Dim itmConfFech     As ListItem
    Dim strSQL          As String
    
    On Error GoTo ErrmPreencheListaConfFech
    
    Screen.MousePointer = vbHourglass
    
    With lvwCadastroConfFech
    
        .ListItems.Clear
        .ColumnHeaders.Clear
        
        'Monta o cabeçalho da lista
        Call mMontaCabecalhoListaConfFech
        
        strSQL = "EXECUTE SP_LISTAR_CONFIG_FECHAMENTO"
        
        Set rsConfFech = New ADODB.Recordset
        rsConfFech.Open strSQL, dbConexaoSMR, adOpenStatic, adLockReadOnly, adCmdText
        
        Do While Not rsConfFech.EOF
        
            Set itmConfFech = .ListItems.Add(, , rsConfFech.Fields("SIGLA_PROJ").Value)
            
            'Segunda
            If rsConfFech.Fields("FLAG_SEG").Value = True Then
                itmConfFech.SubItems(1) = "Sim"
            Else
                itmConfFech.SubItems(1) = "Não"
            End If
            
            itmConfFech.SubItems(2) = rsConfFech.Fields("HR_SEG_INICIO").Value
            itmConfFech.SubItems(3) = rsConfFech.Fields("HR_SEG_FIM").Value
            
            'Terça
            If rsConfFech.Fields("FLAG_TER").Value = True Then
                itmConfFech.SubItems(4) = "Sim"
            Else
                itmConfFech.SubItems(4) = "Não"
            End If
            
            itmConfFech.SubItems(5) = rsConfFech.Fields("HR_TER_INICIO").Value
            itmConfFech.SubItems(6) = rsConfFech.Fields("HR_TER_FIM").Value
            
            'Quarta
            If rsConfFech.Fields("FLAG_QUA").Value = True Then
                itmConfFech.SubItems(7) = "Sim"
            Else
                itmConfFech.SubItems(7) = "Não"
            End If
            
            itmConfFech.SubItems(8) = rsConfFech.Fields("HR_QUA_INICIO").Value
            itmConfFech.SubItems(9) = rsConfFech.Fields("HR_QUA_FIM").Value
            
            'Quinta
            If rsConfFech.Fields("FLAG_QUI").Value = True Then
                itmConfFech.SubItems(10) = "Sim"
            Else
                itmConfFech.SubItems(10) = "Não"
            End If
            
            itmConfFech.SubItems(11) = rsConfFech.Fields("HR_QUI_INICIO").Value
            itmConfFech.SubItems(12) = rsConfFech.Fields("HR_QUI_FIM").Value
            
            'Sexta
            If rsConfFech.Fields("FLAG_SEX").Value = True Then
                itmConfFech.SubItems(13) = "Sim"
            Else
                itmConfFech.SubItems(13) = "Não"
            End If
            
            itmConfFech.SubItems(14) = rsConfFech.Fields("HR_SEX_INICIO").Value
            itmConfFech.SubItems(15) = rsConfFech.Fields("HR_SEX_FIM").Value
            
            'Sabado
            If rsConfFech.Fields("FLAG_SAB").Value = True Then
                itmConfFech.SubItems(16) = "Sim"
            Else
                itmConfFech.SubItems(16) = "Não"
            End If
            
            itmConfFech.SubItems(17) = rsConfFech.Fields("HR_SAB_INICIO").Value
            itmConfFech.SubItems(18) = rsConfFech.Fields("HR_SAB_FIM").Value
            
            'Domingo
            If rsConfFech.Fields("FLAG_DOM").Value = True Then
                itmConfFech.SubItems(19) = "Sim"
            Else
                itmConfFech.SubItems(19) = "Não"
            End If
            
            itmConfFech.SubItems(20) = rsConfFech.Fields("HR_DOM_INICIO").Value
            itmConfFech.SubItems(21) = rsConfFech.Fields("HR_DOM_FIM").Value
            
            rsConfFech.MoveNext
            
        Loop
        
    End With
    
    rsConfFech.Close
    Set rsConfFech = Nothing
    
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrmPreencheListaConfFech:

    Screen.MousePointer = vbNormal
    
    MsgBox "Ocorreu o seguinte erro na rotina mPreencheListaConfFech: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title
           
    If rsConfFech.State = adStateOpen Then
        rsConfFech.Close
        Set rsConfFech = Nothing
    End If

End Sub

Private Sub mTratarTelaAlterarConfFech()

    On Error GoTo ErrmTratarTelaAlterarConfFech
    
    'Desabilita os botões da ToolBar
    Call mDesabilitarBotoesToolBarConfFech
         
    'Coloca o frame que contém os dados sobreposto ao frame de lista
    fraDadosCadConfFech.Visible = True
    fraListaCadConfFech.Visible = False
        
    Exit Sub
    
ErrmTratarTelaAlterarConfFech:

    MsgBox "Ocorreu o seguinte erro na rotina mTratarTelaAlterarConfFech: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title

End Sub


Private Function mValidarCamposConfFech() As Boolean
Dim intHora    As Integer
Dim intMinuto  As Integer
Dim intSegundo As Integer

    On Error GoTo ErrmValidarCamposConfFech
    
    'Segunda
    If chkFechado_Seg.Value = 1 Then
        
        'HR. Inicio
        If Len(Trim(mskInicio_Seg.Text)) <> 8 Then
            MsgBox "HR. Inicio Inválida! Formato hh:mm:ss", vbCritical + vbOKOnly, App.Title
            mskInicio_Seg.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If
        
        intHora = CInt(Mid(mskInicio_Seg.Text, 1, 2))
        intMinuto = CInt(Mid(mskInicio_Seg.Text, 4, 2))
        intSegundo = CInt(Mid(mskInicio_Seg.Text, 7, 2))

        If intHora > 23 Then
            MsgBox "HR. Inicio Inválida!", vbCritical + vbOKOnly, App.Title
            mskInicio_Seg.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If

        If intMinuto > 59 Then
            MsgBox "HR. Inicio Inválida!", vbCritical + vbOKOnly, App.Title
            mskInicio_Seg.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If

        If intSegundo > 59 Then
            MsgBox "HR. Inicio Inválida!", vbCritical + vbOKOnly, App.Title
            mskInicio_Seg.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If

        'HR. Fim
        If Len(Trim(mskFim_Seg.Text)) <> 8 Then
            MsgBox "HR. Fim Inválida! Formato hh:mm:ss", vbCritical + vbOKOnly, App.Title
            mskFim_Seg.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If
        
        intHora = CInt(Mid(mskFim_Seg.Text, 1, 2))
        intMinuto = CInt(Mid(mskFim_Seg.Text, 4, 2))
        intSegundo = CInt(Mid(mskFim_Seg.Text, 7, 2))

        If intHora > 23 Then
            MsgBox "HR. Fim Inválida!", vbCritical + vbOKOnly, App.Title
            mskFim_Seg.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If

        If intMinuto > 59 Then
            MsgBox "HR. Fim Inválida!", vbCritical + vbOKOnly, App.Title
            mskFim_Seg.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If

        If intSegundo > 59 Then
            MsgBox "HR. Fim Inválida!", vbCritical + vbOKOnly, App.Title
            mskFim_Seg.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If
    
    End If
        
    'TERÇA
    If chkFechado_Ter.Value = 1 Then
        
        'HR. Inicio
        If Len(Trim(mskInicio_Ter.Text)) <> 8 Then
            MsgBox "HR. Inicio Inválida! Formato hh:mm:ss", vbCritical + vbOKOnly, App.Title
            mskInicio_Ter.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If
        
        intHora = CInt(Mid(mskInicio_Ter.Text, 1, 2))
        intMinuto = CInt(Mid(mskInicio_Ter.Text, 4, 2))
        intSegundo = CInt(Mid(mskInicio_Ter.Text, 7, 2))

        If intHora > 23 Then
            MsgBox "HR. Inicio Inválida!", vbCritical + vbOKOnly, App.Title
            mskInicio_Ter.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If

        If intMinuto > 59 Then
            MsgBox "HR. Inicio Inválida!", vbCritical + vbOKOnly, App.Title
            mskInicio_Ter.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If

        If intSegundo > 59 Then
            MsgBox "HR. Inicio Inválida!", vbCritical + vbOKOnly, App.Title
            mskInicio_Ter.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If

        'HR. Fim
        If Len(Trim(mskFim_Ter.Text)) <> 8 Then
            MsgBox "HR. Fim Inválida! Formato hh:mm:ss", vbCritical + vbOKOnly, App.Title
            mskFim_Ter.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If
        
        intHora = CInt(Mid(mskFim_Ter.Text, 1, 2))
        intMinuto = CInt(Mid(mskFim_Ter.Text, 4, 2))
        intSegundo = CInt(Mid(mskFim_Ter.Text, 7, 2))

        If intHora > 23 Then
            MsgBox "HR. Fim Inválida!", vbCritical + vbOKOnly, App.Title
            mskFim_Ter.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If

        If intMinuto > 59 Then
            MsgBox "HR. Fim Inválida!", vbCritical + vbOKOnly, App.Title
            mskFim_Ter.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If

        If intSegundo > 59 Then
            MsgBox "HR. Fim Inválida!", vbCritical + vbOKOnly, App.Title
            mskFim_Ter.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If
    
    End If

    'Quarta
    If chkFechado_Qua.Value = 1 Then
        
        'HR. Inicio
        If Len(Trim(mskInicio_Qua.Text)) <> 8 Then
            MsgBox "HR. Inicio Inválida! Formato hh:mm:ss", vbCritical + vbOKOnly, App.Title
            mskInicio_Qua.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If
        
        intHora = CInt(Mid(mskInicio_Qua.Text, 1, 2))
        intMinuto = CInt(Mid(mskInicio_Qua.Text, 4, 2))
        intSegundo = CInt(Mid(mskInicio_Qua.Text, 7, 2))

        If intHora > 23 Then
            MsgBox "HR. Inicio Inválida!", vbCritical + vbOKOnly, App.Title
            mskInicio_Qua.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If

        If intMinuto > 59 Then
            MsgBox "HR. Inicio Inválida!", vbCritical + vbOKOnly, App.Title
            mskInicio_Qua.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If

        If intSegundo > 59 Then
            MsgBox "HR. Inicio Inválida!", vbCritical + vbOKOnly, App.Title
            mskInicio_Qua.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If

        'HR. Fim
        If Len(Trim(mskFim_Qua.Text)) <> 8 Then
            MsgBox "HR. Fim Inválida! Formato hh:mm:ss", vbCritical + vbOKOnly, App.Title
            mskFim_Qua.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If
        
        intHora = CInt(Mid(mskFim_Qua.Text, 1, 2))
        intMinuto = CInt(Mid(mskFim_Qua.Text, 4, 2))
        intSegundo = CInt(Mid(mskFim_Qua.Text, 7, 2))

        If intHora > 23 Then
            MsgBox "HR. Fim Inválida!", vbCritical + vbOKOnly, App.Title
            mskFim_Qua.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If

        If intMinuto > 59 Then
            MsgBox "HR. Fim Inválida!", vbCritical + vbOKOnly, App.Title
            mskFim_Qua.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If

        If intSegundo > 59 Then
            MsgBox "HR. Fim Inválida!", vbCritical + vbOKOnly, App.Title
            mskFim_Qua.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If
    
    End If
        
    'Quinta
    If chkFechado_Qui.Value = 1 Then
        
        'HR. Inicio
        If Len(Trim(mskInicio_Qui.Text)) <> 8 Then
            MsgBox "HR. Inicio Inválida! Formato hh:mm:ss", vbCritical + vbOKOnly, App.Title
            mskInicio_Qui.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If
        
        intHora = CInt(Mid(mskInicio_Qui.Text, 1, 2))
        intMinuto = CInt(Mid(mskInicio_Qui.Text, 4, 2))
        intSegundo = CInt(Mid(mskInicio_Qui.Text, 7, 2))

        If intHora > 23 Then
            MsgBox "HR. Inicio Inválida!", vbCritical + vbOKOnly, App.Title
            mskInicio_Qui.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If

        If intMinuto > 59 Then
            MsgBox "HR. Inicio Inválida!", vbCritical + vbOKOnly, App.Title
            mskInicio_Qui.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If

        If intSegundo > 59 Then
            MsgBox "HR. Inicio Inválida!", vbCritical + vbOKOnly, App.Title
            mskInicio_Qui.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If

        'HR. Fim
        If Len(Trim(mskFim_Qui.Text)) <> 8 Then
            MsgBox "HR. Fim Inválida! Formato hh:mm:ss", vbCritical + vbOKOnly, App.Title
            mskFim_Qui.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If
        
        intHora = CInt(Mid(mskFim_Qui.Text, 1, 2))
        intMinuto = CInt(Mid(mskFim_Qui.Text, 4, 2))
        intSegundo = CInt(Mid(mskFim_Qui.Text, 7, 2))

        If intHora > 23 Then
            MsgBox "HR. Fim Inválida!", vbCritical + vbOKOnly, App.Title
            mskFim_Qui.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If

        If intMinuto > 59 Then
            MsgBox "HR. Fim Inválida!", vbCritical + vbOKOnly, App.Title
            mskFim_Qui.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If

        If intSegundo > 59 Then
            MsgBox "HR. Fim Inválida!", vbCritical + vbOKOnly, App.Title
            mskFim_Qui.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If
    
    End If
        
    'Sexta
    If chkFechado_Sex.Value = 1 Then
        
        'HR. Inicio
        If Len(Trim(mskInicio_Sex.Text)) <> 8 Then
            MsgBox "HR. Inicio Inválida! Formato hh:mm:ss", vbCritical + vbOKOnly, App.Title
            mskInicio_Sex.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If
        
        intHora = CInt(Mid(mskInicio_Sex.Text, 1, 2))
        intMinuto = CInt(Mid(mskInicio_Sex.Text, 4, 2))
        intSegundo = CInt(Mid(mskInicio_Sex.Text, 7, 2))

        If intHora > 23 Then
            MsgBox "HR. Inicio Inválida!", vbCritical + vbOKOnly, App.Title
            mskInicio_Sex.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If

        If intMinuto > 59 Then
            MsgBox "HR. Inicio Inválida!", vbCritical + vbOKOnly, App.Title
            mskInicio_Sex.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If

        If intSegundo > 59 Then
            MsgBox "HR. Inicio Inválida!", vbCritical + vbOKOnly, App.Title
            mskInicio_Sex.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If

        'HR. Fim
        If Len(Trim(mskFim_Sex.Text)) <> 8 Then
            MsgBox "HR. Fim Inválida! Formato hh:mm:ss", vbCritical + vbOKOnly, App.Title
            mskFim_Sex.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If
        
        intHora = CInt(Mid(mskFim_Sex.Text, 1, 2))
        intMinuto = CInt(Mid(mskFim_Sex.Text, 4, 2))
        intSegundo = CInt(Mid(mskFim_Sex.Text, 7, 2))

        If intHora > 23 Then
            MsgBox "HR. Fim Inválida!", vbCritical + vbOKOnly, App.Title
            mskFim_Sex.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If

        If intMinuto > 59 Then
            MsgBox "HR. Fim Inválida!", vbCritical + vbOKOnly, App.Title
            mskFim_Sex.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If

        If intSegundo > 59 Then
            MsgBox "HR. Fim Inválida!", vbCritical + vbOKOnly, App.Title
            mskFim_Sex.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If
    
    End If
        
    'Sabado
    If chkFechado_Sab.Value = 1 Then
        
        'HR. Inicio
        If Len(Trim(mskInicio_Sab.Text)) <> 8 Then
            MsgBox "HR. Inicio Inválida! Formato hh:mm:ss", vbCritical + vbOKOnly, App.Title
            mskInicio_Sab.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If
        
        intHora = CInt(Mid(mskInicio_Sab.Text, 1, 2))
        intMinuto = CInt(Mid(mskInicio_Sab.Text, 4, 2))
        intSegundo = CInt(Mid(mskInicio_Sab.Text, 7, 2))

        If intHora > 23 Then
            MsgBox "HR. Inicio Inválida!", vbCritical + vbOKOnly, App.Title
            mskInicio_Sab.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If

        If intMinuto > 59 Then
            MsgBox "HR. Inicio Inválida!", vbCritical + vbOKOnly, App.Title
            mskInicio_Sab.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If

        If intSegundo > 59 Then
            MsgBox "HR. Inicio Inválida!", vbCritical + vbOKOnly, App.Title
            mskInicio_Sab.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If

        'HR. Fim
        If Len(Trim(mskFim_Sab.Text)) <> 8 Then
            MsgBox "HR. Fim Inválida! Formato hh:mm:ss", vbCritical + vbOKOnly, App.Title
            mskFim_Sab.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If
        
        intHora = CInt(Mid(mskFim_Sab.Text, 1, 2))
        intMinuto = CInt(Mid(mskFim_Sab.Text, 4, 2))
        intSegundo = CInt(Mid(mskFim_Sab.Text, 7, 2))

        If intHora > 23 Then
            MsgBox "HR. Fim Inválida!", vbCritical + vbOKOnly, App.Title
            mskFim_Sab.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If

        If intMinuto > 59 Then
            MsgBox "HR. Fim Inválida!", vbCritical + vbOKOnly, App.Title
            mskFim_Sab.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If

        If intSegundo > 59 Then
            MsgBox "HR. Fim Inválida!", vbCritical + vbOKOnly, App.Title
            mskFim_Sab.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If
    
    End If
        
    'Domingo
    If chkFechado_Dom.Value = 1 Then
        
        'HR. Inicio
        If Len(Trim(mskInicio_Dom.Text)) <> 8 Then
            MsgBox "HR. Inicio Inválida! Formato hh:mm:ss", vbCritical + vbOKOnly, App.Title
            mskInicio_Dom.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If
        
        intHora = CInt(Mid(mskInicio_Dom.Text, 1, 2))
        intMinuto = CInt(Mid(mskInicio_Dom.Text, 4, 2))
        intSegundo = CInt(Mid(mskInicio_Dom.Text, 7, 2))

        If intHora > 23 Then
            MsgBox "HR. Inicio Inválida!", vbCritical + vbOKOnly, App.Title
            mskInicio_Dom.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If

        If intMinuto > 59 Then
            MsgBox "HR. Inicio Inválida!", vbCritical + vbOKOnly, App.Title
            mskInicio_Dom.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If

        If intSegundo > 59 Then
            MsgBox "HR. Inicio Inválida!", vbCritical + vbOKOnly, App.Title
            mskInicio_Dom.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If

        'HR. Fim
        If Len(Trim(mskFim_Dom.Text)) <> 8 Then
            MsgBox "HR. Fim Inválida! Formato hh:mm:ss", vbCritical + vbOKOnly, App.Title
            mskFim_Dom.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If
        
        intHora = CInt(Mid(mskFim_Dom.Text, 1, 2))
        intMinuto = CInt(Mid(mskFim_Dom.Text, 4, 2))
        intSegundo = CInt(Mid(mskFim_Dom.Text, 7, 2))

        If intHora > 23 Then
            MsgBox "HR. Fim Inválida!", vbCritical + vbOKOnly, App.Title
            mskFim_Dom.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If

        If intMinuto > 59 Then
            MsgBox "HR. Fim Inválida!", vbCritical + vbOKOnly, App.Title
            mskFim_Dom.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If

        If intSegundo > 59 Then
            MsgBox "HR. Fim Inválida!", vbCritical + vbOKOnly, App.Title
            mskFim_Dom.SetFocus
            mValidarCamposConfFech = False
            Exit Function
        End If
    
    End If
        
    mValidarCamposConfFech = True
        
    Exit Function
    
ErrmValidarCamposConfFech:

    MsgBox "Ocorreu o seguinte erro na rotina mValidarCamposConfFech: " & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical + vbOKOnly, _
           App.Title
    
End Function

Private Sub chkFechado_Dom_Click()
    Call mClick_Fechamento("Dom")
End Sub

Private Sub chkFechado_Qua_Click()
        Call mClick_Fechamento("Qua")
End Sub

Private Sub chkFechado_Qui_Click()
    Call mClick_Fechamento("Qui")
End Sub

Private Sub chkFechado_Sab_Click()
    Call mClick_Fechamento("Sab")
End Sub

Private Sub chkFechado_Seg_Click()
    Call mClick_Fechamento("Seg")
End Sub

Private Sub chkFechado_Sex_Click()
    Call mClick_Fechamento("Sex")
End Sub

Private Sub chkFechado_Ter_Click()
    Call mClick_Fechamento("Ter")
End Sub

Private Sub cmdCancelar_Click()

    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadConfFech.Visible = False
    fraListaCadConfFech.Visible = True

    'Habilita os botões da ToolBar
    Call mHabilitarBotoesToolBarConfFech

End Sub

Private Sub cmdGravar_Click()
    
    On Error GoTo ErrcmdGravar_Click
    
    If mValidarCamposConfFech() Then
    
        If fraDadosCadConfFech.Tag = "A" Then
            'Alteração
            Call mAlterarConfFech
        End If
        
        'Habilita os botões da ToolBar
        Call mHabilitarBotoesToolBarConfFech
        
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
    Call mHabilitarBotoesToolBarConfFech
        
    'Coloca o frame que contém a lista sobreposto ao frame de dados
    fraDadosCadConfFech.Visible = False
    fraListaCadConfFech.Visible = True

    'Preenche a lista
    Call mPreencheListaConfFech
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmPrincipal.StatusBar.Panels(1).Text = ""
    
End Sub


Private Sub lvwCadastroConfFech_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    With lvwCadastroConfFech
    
        .SortKey = ColumnHeader.Index - 1
        
        If .SortOrder = lvwAscending Then
           .SortOrder = lvwDescending
        Else
           .SortOrder = lvwAscending
        End If
        
        .Sorted = True
        
    End With

End Sub


Private Sub tlbCadastroConfFech_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
       
        Case 2 'Alterar
        
            Call mTratarTelaAlterarConfFech
            Call mDadosConfFech
            Call mHabilitarDesabilitarCamposCadastroConfFech("A")
            
        Case 5 'Sair
            
            Unload Me
            
    End Select

End Sub



Private Sub mClick_Fechamento(strTipo As String)
        
    Select Case strTipo
    
        Case "Seg"   'Segunda
            If chkFechado_Seg.Value = 0 Then
                mskInicio_Seg.Text = "00:00:00"
                mskInicio_Seg.Enabled = False
                mskFim_Seg.Text = "00:00:00"
                mskFim_Seg.Enabled = False
            Else
                mskInicio_Seg.Enabled = True
                mskFim_Seg.Enabled = True
            End If
        Case "Ter"    'Terça
            If chkFechado_Ter.Value = 0 Then
                mskInicio_Ter.Text = "00:00:00"
                mskInicio_Ter.Enabled = False
                mskFim_Ter.Text = "00:00:00"
                mskFim_Ter.Enabled = False
            Else
                mskInicio_Ter.Enabled = True
                mskFim_Ter.Enabled = True
            End If
        Case "Qua"   'Quarta
            If chkFechado_Qua.Value = 0 Then
                mskInicio_Qua.Text = "00:00:00"
                mskInicio_Qua.Enabled = False
                mskFim_Qua.Text = "00:00:00"
                mskFim_Qua.Enabled = False
            Else
                mskInicio_Qua.Enabled = True
                mskFim_Qua.Enabled = True
            End If
        Case "Qui"   'Quinta
            If chkFechado_Qui.Value = 0 Then
                mskInicio_Qui.Text = "00:00:00"
                mskInicio_Qui.Enabled = False
                mskFim_Qui.Text = "00:00:00"
                mskFim_Qui.Enabled = False
            Else
                mskInicio_Qui.Enabled = True
                mskFim_Qui.Enabled = True
            End If
        Case "Sex"   'Sexta
            If chkFechado_Sex.Value = 0 Then
                mskInicio_Sex.Text = "00:00:00"
                mskInicio_Sex.Enabled = False
                mskFim_Sex.Text = "00:00:00"
                mskFim_Sex.Enabled = False
            Else
                mskInicio_Sex.Enabled = True
                mskFim_Sex.Enabled = True
            End If
        Case "Sab"   'Sabado
            If chkFechado_Sab.Value = 0 Then
                mskInicio_Sab.Text = "00:00:00"
                mskInicio_Sab.Enabled = False
                mskFim_Sab.Text = "00:00:00"
                mskFim_Sab.Enabled = False
            Else
                mskInicio_Sab.Enabled = True
                mskFim_Sab.Enabled = True
            End If
        Case "Dom"   'Domingo
            If chkFechado_Dom.Value = 0 Then
                mskInicio_Dom.Text = "00:00:00"
                mskInicio_Dom.Enabled = False
                mskFim_Dom.Text = "00:00:00"
                mskFim_Dom.Enabled = False
            Else
                mskInicio_Dom.Enabled = True
                mskFim_Dom.Enabled = True
            End If
    End Select

End Sub
