VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAtualizarCases_TesteInt_Sin 
   Caption         =   "Atualização dos Flag Teste Integrado -> Sinergia"
   ClientHeight    =   1650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6060
   Icon            =   "frmAtualizarCases_TesteInt_Sin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraAtualizacaoTar 
      Height          =   750
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   5775
      Begin MSComctlLib.ProgressBar prbExp 
         Height          =   360
         Left            =   105
         TabIndex        =   3
         Top             =   270
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.CommandButton cmdCancela 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4545
      TabIndex        =   1
      Top             =   1170
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3225
      TabIndex        =   0
      Top             =   1185
      Width           =   1215
   End
End
Attribute VB_Name = "frmAtualizarCases_TesteInt_Sin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()

    cmdOK.Enabled = False
    cmdCancela.Enabled = False
    
    fraAtualizacaoTar.Caption = ""
    
    If Atualizar_Cases_TesteInt_Sin() Then
        MsgBox "Atualização dos Flag Teste Integrado -> Sinergia gerada com sucesso.", vbInformation + vbOKOnly, "Atualização dos Flag Teste Integrado -> Sinergia"
    Else
        MsgBox "Erro na Atualização dos Flag Teste Integrado -> Sinergia.", vbCritical + vbOKOnly, "Atualização dos Flag Teste Integrado -> Sinergia"
    End If
    
    cmdOK.Enabled = True
    cmdCancela.Enabled = True
    prbExp.Value = 0
    fraAtualizacaoTar.Caption = ""
    
End Sub

Private Sub cmdCancela_Click()
    Unload Me
End Sub

Private Function Atualizar_Cases_TesteInt_Sin() As Boolean

Dim rs1               As New ADODB.Recordset
Dim rs2               As New ADODB.Recordset
Dim rs3               As New ADODB.Recordset
Dim cmdAtualizar      As New ADODB.Command

Dim strSQL            As String

On Error GoTo ErrAtualizar_Cases_TesteInt_Sin

        fraAtualizacaoTar.Caption = "Atualização dos Flag Teste Integrado -> Sinergia"
        
        Atualizar_Cases_TesteInt_Sin = False
        
        Screen.MousePointer = vbHourglass
        strSQL = "SP_EXCLUIR_FLAG_10_DATAS_FIM_7"
        
        dbConexaoSMR.Execute strSQL
        
        
        strSQL = "SP_LISTAR_ATUALIZA_CASES_TESTEINT_SIN "
                    
        rs1.CursorLocation = adUseClient
        rs2.CursorLocation = adUseClient
        rs3.CursorLocation = adUseClient
        
        rs1.Open strSQL, dbConexaoTIN, adOpenStatic, adLockReadOnly, adCmdText
        
        prbExp.Max = rs1.RecordCount
        prbExp.Min = 0
        prbExp.Value = 0
        
        Do While Not rs1.EOF
        
            strSQL = "SP_LISTAR_TAREFAS_DESENV_TIN_SIN " & Trim(Str(rs1("Cod_Proj_Orig"))) & ", " & Trim(Str(rs1("Cod_UID_Orig")))
            
            rs2.Open strSQL, dbConexaoTIN, adOpenStatic, adLockReadOnly, adCmdText
            
            If Not rs2.EOF Then
                
                strSQL = "SP_LISTAR_TAREFAS_DESENV_SIN_TIN " & Trim(Str(rs1("Cod_Proj_Dest"))) & ", " & Trim(Str(rs1("Cod_UID_Dest")))
                
                rs3.Open strSQL, dbConexaoSMR, adOpenStatic, adLockReadOnly, adCmdText
                
                If Not rs3.EOF Then
                        
                    If rs2("FLAG_10") = True Then
                    
                        With cmdAtualizar
                        
                            .ActiveConnection = dbConexaoSMR
                            .CommandType = 4
                            .CommandTimeout = 1200
                            .CommandText = "SP_ATUALIZAR_CASES_TESTEINT_SIN"
                            
                            .Parameters.Refresh
                            .Parameters(1).Value = rs3("PROJ_ID")
                            .Parameters(2).Value = rs3("TASK_UID")
                            .Parameters(3).Value = rs2("DT_FIM_7")
                        
                        End With
                        
                        cmdAtualizar.Execute
                        
                    End If
                End If
            End If
            
            rs2.Close
            rs3.Close
            
            prbExp.Value = prbExp.Value + 1
            
            rs1.MoveNext
        
        Loop
        
        rs1.Close
        
        Screen.MousePointer = vbNormal

        Atualizar_Cases_TesteInt_Sin = True
        
        Exit Function
        
ErrAtualizar_Cases_TesteInt_Sin:
    
    Screen.MousePointer = vbNormal
    Atualizar_Cases_TesteInt_Sin = False
    
End Function

Private Sub Form_Unload(Cancel As Integer)

    frmPrincipal.StatusBar.Panels(1).Text = ""
    
End Sub

